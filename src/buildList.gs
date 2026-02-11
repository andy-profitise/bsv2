/************************************************************
 * BUILD LIST v2 - Monday.com Only + Two-Account Inbox
 *
 * Vendors sourced ENTIRELY from monday.com boards (no L1M/L6M).
 * Hot zone based on current user's inbox + email log (captures
 * the second user's emails from shared log).
 *
 * Sort: Hot zone on top, then by status priority, then alpha.
 *
 * Required sheets:
 *   - "buyers monday.com" (synced via Sync monday.com Data)
 *   - "affiliates monday.com" (synced via Sync monday.com Data)
 *   - "Settings" (blacklist col E, sublabel map cols A-B, label config cols G-I)
 *
 * Output: "List" sheet with columns:
 *   Vendor | Source | Status | Notes | Gmail Link | no snoozing | processed?
 ************************************************************/

/***** CONFIG *****/
const SHEET_MON_BUYERS       = 'buyers monday.com';
const SHEET_MON_AFFILIATES   = 'affiliates monday.com';

const SHEET_OUT              = 'List';
const SHEET_SETTINGS         = 'Settings';

// monday.com special columns (0-based indexes)
const ALIAS_COL_INDEX = 15;           // Column P (1-based 16) -> 0-based 15
const BUYERS_STATUS_INDEX = 18;       // Column S (1-based 19) -> 0-based 18
const AFFILIATES_STATUS_INDEX = 18;   // Column S (1-based 19) -> 0-based 18
const BUYERS_NOTES_INDEX = 3;         // Column D (1-based 4) -> 0-based 3
const AFFILIATES_NOTES_INDEX = 3;     // Column D (1-based 4) -> 0-based 3

// Status priority for sorting
const STATUS_PRIORITY = [
  'live',
  'onboarding',
  'paused',
  'preonboarding',
  'early talks',
  'top 500 remodelers',
  'other',
  'dead'
];
const STATUS_RANK = STATUS_PRIORITY.reduce((m, s, i) => (m[s] = i, m), {});

// Skip reason tracking (for debugging)
const SKIP_REASONS = [];


/**
 * Build the vendor List from monday.com boards.
 * Hot zone: vendors with emails in current user's inbox OR in the shared email log.
 * Normal zone: all other vendors sorted by status priority then alphabetically.
 *
 * Output columns: Vendor | Source | Status | Notes | Gmail Link | no snoozing | processed?
 */
function buildListWithGmailAndNotes() {
  const ss = SpreadsheetApp.getActive();

  console.log('=== BUILD LIST (monday.com only) START ===');

  const shMonB = ss.getSheetByName(SHEET_MON_BUYERS);
  const shMonA = ss.getSheetByName(SHEET_MON_AFFILIATES);

  if (!shMonB && !shMonA) {
    SpreadsheetApp.getUi().alert(
      'No monday.com data sheets found.\n\n' +
      'Run "Sync monday.com Data" first to create the\n' +
      '"buyers monday.com" and "affiliates monday.com" sheets.'
    );
    return;
  }

  const blacklist = readBlacklist_(ss);

  // Read vendors from monday.com boards (preserving board order via rowIndex)
  const existingSet = new Set();
  const buyers = shMonB ? readMondaySheet_(shMonB, 'Buyer', 'buyers monday.com', blacklist, existingSet) : [];
  const affiliates = shMonA ? readMondaySheet_(shMonA, 'Affiliate', 'affiliates monday.com', blacklist, existingSet) : [];

  // Dedupe across both boards (buyers win on collision)
  const all = dedupeKeepFirst_([...buyers, ...affiliates]);

  console.log('Vendors loaded:', { buyers: buyers.length, affiliates: affiliates.length, total: all.length });

  // Sort by status priority, then type (buyers first), then alphabetical
  all.sort((a, b) => {
    const sA = STATUS_RANK[String(a.status || '').toLowerCase()] ?? STATUS_RANK['other'];
    const sB = STATUS_RANK[String(b.status || '').toLowerCase()] ?? STATUS_RANK['other'];
    if (sA !== sB) return sA - sB;
    const rankA = (a.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    const rankB = (b.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    if (rankA !== rankB) return rankA - rankB;
    return String(a.name).localeCompare(String(b.name));
  });

  // Fetch contact emails for hot zone detection (reused later for Gmail links)
  ss.toast('Fetching contact emails for hot zone...', '📇 Contacts', 10);
  const contactEmailMapForHot = fetchAllContactEmails_();
  const vendorContactEmailMapForHot = buildVendorContactEmailMap_(shMonB, shMonA, contactEmailMapForHot);

  // HOT ZONE: Detect vendors with inbox emails (current user + email log)
  console.log('Detecting hot vendors...');
  const hotVendorSet = getHotVendors_(all, vendorContactEmailMapForHot);
  console.log('Hot vendors found:', hotVendorSet.size);

  const hotZone = [];
  const normalZone = [];

  for (const r of all) {
    if (hotVendorSet.has(r.name.toLowerCase())) {
      hotZone.push(r);
    } else {
      normalZone.push(r);
    }
  }

  // Final list: Hot zone at top (keeps status sort within), then normal zone
  const finalList = [...hotZone, ...normalZone];

  console.log('Total vendors for output:', finalList.length);

  // Write to List sheet
  const shOut = ensureSheet_(ss, SHEET_OUT);

  const lastRow = shOut.getMaxRows();
  const lastCol = shOut.getMaxColumns();
  if (lastRow > 0 && lastCol > 0) {
    shOut.getRange(1, 1, lastRow, lastCol).clear();
  }

  // Headers: dual Gmail link columns for both team members
  const team = getTeamMembers_();
  shOut.getRange(1, 1, 1, 9).setValues([[
    'Vendor', 'Source', 'Status', 'Notes',
    `${team.m1.name} Gmail`, `${team.m1.name} no snooze`,
    `${team.m2.name} Gmail`, `${team.m2.name} no snooze`,
    'processed?'
  ]]).setFontWeight('bold').setBackground('#e8f0fe');

  if (finalList.length > 0) {
    // Write data (columns A-D)
    const data = finalList.map(r => [
      r.name,
      r.source,
      r.status || '',
      r.notes || ''
    ]);

    shOut.getRange(2, 1, data.length, 4).setValues(data);

    // Read Gmail sublabel mappings from Settings sheet
    const gmailSublabelMap = readGmailSublabelMap_(ss);

    // Reuse contact email map already fetched for hot zone detection
    // Build Gmail links for both team members using contact emails
    const gmailData = finalList.map(v => {
      const vendorKey = v.name.toLowerCase();
      const vendorSlug = gmailSublabelMap.has(vendorKey)
        ? gmailSublabelMap.get(vendorKey)
        : null;

      // Get contact emails for this vendor
      const contactEmails = vendorContactEmailMapForHot.get(vendorKey) || [];

      const queries = buildVendorEmailQuery_(v.name, vendorSlug, contactEmails);
      const gmailAll = buildGmailSearchUrl_(queries.allQuery);
      const gmailNoSnooze = buildGmailSearchUrl_(queries.noSnoozeQuery);

      // Same search URL for both people (each opens in their own /u/0/ primary account)
      return [gmailAll, gmailNoSnooze, gmailAll, gmailNoSnooze, false];
    });

    shOut.getRange(2, 5, gmailData.length, 5).setValues(gmailData);

    // Highlight hot zone rows
    if (hotZone.length > 0) {
      shOut.getRange(2, 1, hotZone.length, 9).setBackground('#e8f5e9'); // Light green
    }
  }

  // Auto-resize columns
  shOut.autoResizeColumns(1, 4);
  shOut.setColumnWidth(5, 110);
  shOut.setColumnWidth(6, 110);
  shOut.setColumnWidth(7, 110);
  shOut.setColumnWidth(8, 110);
  shOut.setColumnWidth(9, 100);

  console.log('=== BUILD LIST END ===');

  const counts = {
    'HOT (inbox emails)': hotZone.length,
    'Total vendors': finalList.length,
    'Buyers': finalList.filter(v => (v.type || '').toLowerCase().startsWith('buyer')).length,
    'Affiliates': finalList.filter(v => (v.type || '').toLowerCase().startsWith('affiliate')).length
  };

  console.log('Final counts:', counts);

  ss.toast(
    Object.entries(counts).map(([k,v]) => `${k}: ${v}`).join(' | '),
    'List Built',
    8
  );
}


/** ========== HOT ZONE DETECTION (Two Accounts) ========== **/

/**
 * Detect "hot" vendors that have emails in either:
 * 1. Current user's Gmail inbox (live search)
 * 2. Email log spreadsheet (captures both users' emails)
 *
 * The email log is the key to two-account support: each user runs
 * "Scan Inbox to Log" which records their inbox emails. The hot zone
 * detection then reads from the shared log to see both accounts.
 *
 * @param {Array} allVendors - Array of vendor objects
 * @param {Map} [vendorContactEmailMap] - Map of lowercased vendor name → email addresses
 * @returns {Set} Set of lowercased vendor names that are "hot"
 */
function getHotVendors_(allVendors, vendorContactEmailMap) {
  const hotSet = new Set();
  const cfg = getLabelConfig_();

  // Build vendor name lookup for fast matching
  const vendorMap = new Map();
  const vendorNames = allVendors.map(v => {
    const nameLower = v.name.toLowerCase();
    vendorMap.set(nameLower, v.name);
    return { name: v.name, nameLower: nameLower };
  });

  // --- SOURCE 1: Current user's Gmail inbox (live) ---
  try {
    const liveQuery = cfg.live_email_query || 'label:inbox';
    console.log(`Live Gmail search: ${liveQuery}`);
    const threads = GmailApp.search(liveQuery, 0, 200);
    console.log(`Found ${threads.length} threads in current user inbox`);

    const vendorLabelPrefix = cfg.vendor_label_prefix || '';
    const vendorLabelStyle = cfg.vendor_label_style || 'none';

    // Build reverse map: email address → vendor name (for contact-based matching)
    const emailToVendorMap = new Map();
    if (vendorContactEmailMap) {
      for (const [vendorKey, emails] of vendorContactEmailMap) {
        for (const email of emails) {
          emailToVendorMap.set(email.toLowerCase(), vendorKey);
        }
      }
    }

    for (const thread of threads) {
      try {
        matchThreadToVendor_(thread, vendorNames, vendorMap, vendorLabelPrefix, vendorLabelStyle, hotSet, emailToVendorMap);
      } catch (e) {
        console.log(`Error processing thread: ${e.message}`);
      }
    }

    console.log(`After live inbox: ${hotSet.size} hot vendors`);
  } catch (e) {
    console.log(`Error searching current user Gmail: ${e.message}`);
  }

  // --- SOURCE 2: Email log spreadsheet (both accounts) ---
  try {
    const logHotVendors = getHotVendorsFromLog_(vendorNames);
    for (const name of logHotVendors) {
      hotSet.add(name);
    }
    console.log(`After email log: ${hotSet.size} hot vendors (+${logHotVendors.size} from log)`);
  } catch (e) {
    console.log(`Error reading email log: ${e.message}`);
  }

  return hotSet;
}

/**
 * Try to match a Gmail thread to a vendor name.
 * Uses contact email matching first, then vendor labels, then name matching.
 *
 * @param {GmailThread} thread
 * @param {Array} vendorNames - [{name, nameLower}, ...]
 * @param {Map} vendorMap - lowercased name → original name
 * @param {string} vendorLabelPrefix
 * @param {string} vendorLabelStyle
 * @param {Set} hotSet - Set to add matched vendor names to
 * @param {Map} [emailToVendorMap] - email address → lowercased vendor name
 */
function matchThreadToVendor_(thread, vendorNames, vendorMap, vendorLabelPrefix, vendorLabelStyle, hotSet, emailToVendorMap) {

  // METHOD 1: Match by contact email addresses (most reliable)
  if (emailToVendorMap && emailToVendorMap.size > 0) {
    const messages = thread.getMessages();
    if (messages.length > 0) {
      for (const msg of messages) {
        const participants = [
          msg.getFrom() || '',
          msg.getTo() || '',
          msg.getCc() || ''
        ].join(' ').toLowerCase();

        for (const [contactEmail, vendorKey] of emailToVendorMap) {
          if (participants.includes(contactEmail)) {
            hotSet.add(vendorKey);
            return;
          }
        }
      }
    }
  }

  // METHOD 2: Check vendor labels (if configured)
  if (vendorLabelStyle !== 'none' && vendorLabelPrefix) {
    const labels = thread.getLabels();
    for (const label of labels) {
      const labelName = label.getName();

      if (vendorLabelStyle === 'sublabel' && labelName.startsWith(vendorLabelPrefix)) {
        const vendorNameFromLabel = labelName.substring(vendorLabelPrefix.length).toLowerCase();
        if (vendorMap.has(vendorNameFromLabel)) {
          hotSet.add(vendorNameFromLabel);
          return;
        }
        for (const vendor of vendorNames) {
          if (vendorNameFromLabel.includes(vendor.nameLower) || vendor.nameLower.includes(vendorNameFromLabel)) {
            hotSet.add(vendor.nameLower);
            return;
          }
        }
      } else if (vendorLabelStyle === 'flat' && labelName.toLowerCase().startsWith(vendorLabelPrefix.toLowerCase())) {
        const slugFromLabel = labelName.toLowerCase().substring(vendorLabelPrefix.length);
        for (const vendor of vendorNames) {
          const vendorSlug = vendor.nameLower.replace(/[^a-z0-9-]+/g, '-').replace(/^-+|-+$/g, '');
          if (slugFromLabel === vendorSlug) {
            hotSet.add(vendor.nameLower);
            return;
          }
        }
      }
    }
  }

  // METHOD 3: Name match in subject/sender/recipient (fallback)
  const subject = (thread.getFirstMessageSubject() || '').toLowerCase();
  const messages2 = thread.getMessages();
  let emailText = subject;
  if (messages2.length > 0) {
    const firstMsg = messages2[0];
    emailText += ' ' + (firstMsg.getFrom() || '').toLowerCase();
    emailText += ' ' + (firstMsg.getTo() || '').toLowerCase();
  }

  for (const vendor of vendorNames) {
    if (vendor.nameLower.length >= 3 && emailText.includes(vendor.nameLower)) {
      hotSet.add(vendor.nameLower);
      return;
    }
  }
}

/**
 * Get hot vendors from the shared email log.
 * Looks at emails logged within the last 7 days from ANY user.
 * This is how the second person's inbox gets captured.
 *
 * @param {Array} vendorNames - Array of {name, nameLower} objects
 * @returns {Set} Set of lowercased vendor names found in email log
 */
function getHotVendorsFromLog_(vendorNames) {
  const hotFromLog = new Set();

  let logSS;
  try {
    logSS = getEmailLogSpreadsheet_();
    if (!logSS) return hotFromLog;
  } catch (e) {
    return hotFromLog;
  }

  const logSheet = logSS.getSheetByName('Email Log');
  if (!logSheet) return hotFromLog;

  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return hotFromLog;

  // Read the log: columns A (timestamp), K (vendor), H (date)
  const data = logSheet.getRange(2, 1, lastRow - 1, 14).getValues();

  // Only consider emails logged in the last 7 days
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);

  for (const row of data) {
    const timestamp = new Date(row[0]);  // Column A: Timestamp
    const vendor = String(row[10] || ''); // Column K: Vendor

    if (timestamp >= cutoff && vendor) {
      hotFromLog.add(vendor.toLowerCase());
    }
  }

  console.log(`Email log: ${hotFromLog.size} vendors with emails in last 7 days`);
  return hotFromLog;
}

/**
 * Menu entry: Scan current user's inbox and log all emails.
 * Each team member should run this periodically so their emails
 * appear in the shared email log for hot zone detection.
 */
function scanInboxToLog() {
  const ss = SpreadsheetApp.getActive();
  const cfg = getLabelConfig_();
  const liveQuery = cfg.live_email_query || 'label:inbox';

  ss.toast('Scanning inbox...', 'Email Log', 5);

  const threads = GmailApp.search(liveQuery, 0, 200);
  console.log(`Scanning ${threads.length} inbox threads for logging`);

  // Get vendor list for matching
  const listSh = ss.getSheetByName(SHEET_OUT);
  const vendorNames = [];

  if (listSh && listSh.getLastRow() > 1) {
    const data = listSh.getRange(2, 1, listSh.getLastRow() - 1, 1).getValues().flat();
    for (const name of data) {
      const n = String(name || '').trim();
      if (n) vendorNames.push(n);
    }
  }

  // Fetch contact emails for vendor matching
  const shMonB = ss.getSheetByName(SHEET_MON_BUYERS);
  const shMonA = ss.getSheetByName(SHEET_MON_AFFILIATES);
  const contactEmailMap = fetchAllContactEmails_();
  const vendorEmailMap = buildVendorContactEmailMap_(shMonB, shMonA, contactEmailMap);

  let totalLogged = 0;

  for (const thread of threads) {
    try {
      const messages = thread.getMessages();
      if (messages.length === 0) continue;

      const lastMessage = messages[messages.length - 1];
      const subject = thread.getFirstMessageSubject() || '(no subject)';
      const date = lastMessage.getDate();
      const labels = thread.getLabels().map(l => l.getName()).join(', ');
      const threadId = thread.getId();

      const tz = 'America/Los_Angeles';
      const dateFormatted = Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm');

      let snippet = '';
      try {
        snippet = lastMessage.getPlainBody().substring(0, 200) + '...';
      } catch (e) {
        snippet = '';
      }

      // Match to a vendor — prefer contact email matching, fall back to name matching
      const searchText = (subject + ' ' + (lastMessage.getFrom() || '') + ' ' + (lastMessage.getTo() || '') + ' ' + (lastMessage.getCc() || '')).toLowerCase();
      let matchedVendor = '';

      // Try contact email matching first
      if (vendorEmailMap && vendorEmailMap.size > 0) {
        for (const [vendorKey, emails] of vendorEmailMap) {
          for (const email of emails) {
            if (searchText.includes(email.toLowerCase())) {
              matchedVendor = vendorNames.find(n => n.toLowerCase() === vendorKey) || vendorKey;
              break;
            }
          }
          if (matchedVendor) break;
        }
      }

      // Fall back to name matching
      if (!matchedVendor) {
        for (const vName of vendorNames) {
          if (vName.length >= 3 && searchText.includes(vName.toLowerCase())) {
            matchedVendor = vName;
            break;
          }
        }
      }

      const emailObj = {
        threadId: threadId,
        subject: subject,
        date: dateFormatted,
        count: messages.length,
        labels: labels,
        link: buildGmailThreadUrl_(threadId),
        snippet: snippet,
        isSnoozed: false
      };

      logEmailsToSheet_([emailObj], matchedVendor || '(unmatched)', [thread]);
      totalLogged++;

    } catch (e) {
      console.log(`Error logging thread: ${e.message}`);
    }
  }

  ss.toast(`Logged ${totalLogged} email threads to shared log`, 'Inbox Scan Complete', 5);
}


/** ========== MONDAY.COM SHEET READER ========== **/

/**
 * Read vendors from a monday.com sheet.
 * Preserves board row order (rowIndex) for ordering reference.
 *
 * Returns: [{name, type, source, status, notes, rowIndex}, ...]
 */
function readMondaySheet_(sh, type, sourceLabel, blacklist, existingSet) {
  console.log(`\n[${sourceLabel}] START READ`);

  const allValues = sh.getDataRange().getValues();
  if (allValues.length < 2) return [];

  const headers = allValues[0].map(h => String(h || '').trim());
  console.log(`[${sourceLabel}] Total rows:`, allValues.length - 1);

  const nameIdx = 0;
  const typeIdx = headers.findIndex(h => eq_(h, 'Type'));
  const statusIdx = type === 'Buyer' ? BUYERS_STATUS_INDEX : AFFILIATES_STATUS_INDEX;
  const notesIdx = type === 'Buyer' ? BUYERS_NOTES_INDEX : AFFILIATES_NOTES_INDEX;

  const out = [];
  let skippedBlacklist = 0;
  let skippedExisting = 0;
  let skippedAlias = 0;
  let skippedNoName = 0;

  for (let i = 1; i < allValues.length; i++) {
    const row = allValues[i];
    const rawName = row[nameIdx] || '';
    const name = normalizeName_(rawName);

    if (!name) { skippedNoName++; continue; }

    const key = name.toLowerCase();
    if (blacklist.has(key)) { skippedBlacklist++; continue; }
    if (existingSet.has(key)) { skippedExisting++; continue; }

    // Aliases in Column P (index 15)
    const aliasRaw = (ALIAS_COL_INDEX < row.length) ? (row[ALIAS_COL_INDEX] || '') : '';
    const aliases = String(aliasRaw).split(',')
      .map(s => normalizeName_(s).toLowerCase())
      .filter(Boolean);

    if (aliases.some(a => existingSet.has(a) || blacklist.has(a))) {
      skippedAlias++;
      SKIP_REASONS.push({ name, source: sourceLabel, reason: 'alias-duplicate' });
      continue;
    }

    const explicitType = (typeIdx !== -1 && typeIdx < row.length) ? String(row[typeIdx] || '').trim() : '';
    const finalType = explicitType || type;
    const statusRaw = (statusIdx != null && statusIdx < row.length) ? String(row[statusIdx] || '').trim() : '';
    const status = normalizeStatus_(statusRaw);
    const notes = (notesIdx != null && notesIdx < row.length) ? String(row[notesIdx] || '').trim() : '';

    out.push({ name, type: finalType, source: sourceLabel, status, notes, rowIndex: i });
    existingSet.add(key);
  }

  console.log(`[${sourceLabel}] Skip summary:`, {
    noName: skippedNoName, blacklist: skippedBlacklist,
    existing: skippedExisting, alias: skippedAlias, added: out.length
  });
  console.log(`[${sourceLabel}] END READ\n`);

  return out;
}


/** ========== STATUS & NOTES LOOKUP ========== **/

function buildStatusMaps_(shMonB, shMonA) {
  const buyersMap = new Map();
  const affiliatesMap = new Map();

  if (shMonB) {
    const data = shMonB.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const vendor = normalizeName_(data[i][0]);
      const status = BUYERS_STATUS_INDEX < data[i].length ? String(data[i][BUYERS_STATUS_INDEX] || '').trim() : '';
      if (vendor) buyersMap.set(vendor.toLowerCase(), status);
    }
  }
  if (shMonA) {
    const data = shMonA.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const vendor = normalizeName_(data[i][0]);
      const status = AFFILIATES_STATUS_INDEX < data[i].length ? String(data[i][AFFILIATES_STATUS_INDEX] || '').trim() : '';
      if (vendor) affiliatesMap.set(vendor.toLowerCase(), status);
    }
  }
  return { buyers: buyersMap, affiliates: affiliatesMap };
}

function buildNotesMaps_(shMonB, shMonA) {
  const buyersMap = new Map();
  const affiliatesMap = new Map();

  if (shMonB) {
    const data = shMonB.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const vendor = normalizeName_(data[i][0]);
      const notes = BUYERS_NOTES_INDEX < data[i].length ? String(data[i][BUYERS_NOTES_INDEX] || '').trim() : '';
      if (vendor) buyersMap.set(vendor.toLowerCase(), notes);
    }
  }
  if (shMonA) {
    const data = shMonA.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const vendor = normalizeName_(data[i][0]);
      const notes = AFFILIATES_NOTES_INDEX < data[i].length ? String(data[i][AFFILIATES_NOTES_INDEX] || '').trim() : '';
      if (vendor) affiliatesMap.set(vendor.toLowerCase(), notes);
    }
  }
  return { buyers: buyersMap, affiliates: affiliatesMap };
}

function lookupStatus_(name, type, statusMaps) {
  const key = name.toLowerCase();
  const isBuyer = (type || '').toLowerCase().startsWith('buyer');
  if (isBuyer && statusMaps.buyers.has(key)) return normalizeStatus_(statusMaps.buyers.get(key));
  if (!isBuyer && statusMaps.affiliates.has(key)) return normalizeStatus_(statusMaps.affiliates.get(key));
  if (statusMaps.buyers.has(key)) return normalizeStatus_(statusMaps.buyers.get(key));
  if (statusMaps.affiliates.has(key)) return normalizeStatus_(statusMaps.affiliates.get(key));
  return '';
}

function lookupNotes_(name, type, notesMaps) {
  const key = name.toLowerCase();
  const isBuyer = (type || '').toLowerCase().startsWith('buyer');
  if (isBuyer && notesMaps.buyers.has(key)) return notesMaps.buyers.get(key);
  if (!isBuyer && notesMaps.affiliates.has(key)) return notesMaps.affiliates.get(key);
  if (notesMaps.buyers.has(key)) return notesMaps.buyers.get(key);
  if (notesMaps.affiliates.has(key)) return notesMaps.affiliates.get(key);
  return '';
}


/** ========== GMAIL SUBLABEL MAPPING ========== **/

function readGmailSublabelMap_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const map = new Map();
  if (!sh) return map;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return map;

  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const row of data) {
    const vendorName = String(row[0] || '').trim();
    const sublabel = String(row[1] || '').trim();
    if (vendorName && sublabel) map.set(vendorName.toLowerCase(), sublabel);
  }
  console.log(`Loaded ${map.size} Gmail sublabel mappings from Settings`);
  return map;
}


/** ========== CONTACT EMAIL LOOKUP ========== **/

// Column indexes for the "Contacts" column in synced monday.com sheets
const BUYERS_CONTACTS_COL_INDEX = 7;       // "Contacts" column in buyers monday.com
const AFFILIATES_CONTACTS_COL_INDEX = 4;   // "Contacts" column in affiliates monday.com

/**
 * Fetch ALL contacts from the monday.com Contacts board with their email addresses.
 * Returns a Map<lowercased contact name, email address>.
 *
 * This is a single batch call — much faster than per-vendor lookups.
 */
function fetchAllContactEmails_() {
  const apiToken = PropertiesService.getScriptProperties().getProperty('MONDAY_API_TOKEN') || '';
  if (!apiToken) {
    console.log('No MONDAY_API_TOKEN set — skipping contact email fetch');
    return new Map();
  }

  const contactsBoardId = '9304296922'; // BS_CFG.CONTACTS_BOARD_ID
  const emailColumnId = 'email_mkrk53z4';
  const contactMap = new Map(); // lowercased name → email

  let cursor = null;
  let pageCount = 0;
  const maxPages = 20;

  do {
    pageCount++;
    const cursorPart = cursor ? `, cursor: "${cursor}"` : '';
    const query = `
      query {
        boards(ids: [${contactsBoardId}]) {
          items_page(limit: 500${cursorPart}) {
            cursor
            items {
              name
              column_values(ids: ["${emailColumnId}"]) {
                id
                text
              }
            }
          }
        }
      }
    `;

    try {
      const response = UrlFetchApp.fetch('https://api.monday.com/v2', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': apiToken },
        payload: JSON.stringify({ query }),
        muteHttpExceptions: true
      });

      const result = JSON.parse(response.getContentText());
      const itemsPage = result.data?.boards?.[0]?.items_page;
      if (!itemsPage) break;

      for (const item of itemsPage.items) {
        const name = String(item.name || '').trim();
        const emailCol = item.column_values.find(cv => cv.id === emailColumnId);
        const email = String(emailCol?.text || '').trim();

        if (name && email && email.includes('@')) {
          contactMap.set(name.toLowerCase(), email);
        }
      }

      cursor = itemsPage.cursor;
    } catch (e) {
      console.log(`Error fetching contacts page ${pageCount}: ${e.message}`);
      break;
    }
  } while (cursor && pageCount < maxPages);

  console.log(`Fetched ${contactMap.size} contacts with email addresses from Contacts board`);
  return contactMap;
}

/**
 * Read vendor → contact names from the synced monday.com sheets.
 * The "Contacts" column contains comma-separated display names of linked contacts.
 *
 * @param {Sheet} sh - The synced monday.com sheet
 * @param {number} contactsColIdx - 0-based column index for the Contacts column
 * @returns {Map<string, string[]>} Map of lowercased vendor name → array of lowercased contact names
 */
function readVendorContactNames_(sh, contactsColIdx) {
  const map = new Map();
  if (!sh) return map;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return map;

  for (let i = 1; i < data.length; i++) {
    const vendor = normalizeName_(data[i][0]);
    if (!vendor) continue;

    const contactsRaw = String(data[i][contactsColIdx] || '').trim();
    if (!contactsRaw) continue;

    // display_value is comma-separated contact names
    const contactNames = contactsRaw.split(',')
      .map(n => n.trim().toLowerCase())
      .filter(Boolean);

    if (contactNames.length > 0) {
      map.set(vendor.toLowerCase(), contactNames);
    }
  }

  return map;
}

/**
 * Build a map of vendor name → contact email addresses.
 * Combines the synced contact names from monday.com sheets with
 * the batch-fetched email addresses from the Contacts board.
 *
 * @param {Sheet} shMonB - Buyers monday.com sheet
 * @param {Sheet} shMonA - Affiliates monday.com sheet
 * @param {Map} contactEmailMap - Map from fetchAllContactEmails_()
 * @returns {Map<string, string[]>} Map of lowercased vendor name → array of email addresses
 */
function buildVendorContactEmailMap_(shMonB, shMonA, contactEmailMap) {
  const vendorEmailMap = new Map();

  // Read vendor→contact name mappings from both boards
  const buyerContacts = readVendorContactNames_(shMonB, BUYERS_CONTACTS_COL_INDEX);
  const affiliateContacts = readVendorContactNames_(shMonA, AFFILIATES_CONTACTS_COL_INDEX);

  // Merge both into a single map
  const allVendorContacts = new Map([...buyerContacts, ...affiliateContacts]);

  // Resolve contact names to email addresses
  for (const [vendorKey, contactNames] of allVendorContacts) {
    const emails = [];
    for (const contactName of contactNames) {
      const email = contactEmailMap.get(contactName);
      if (email) {
        emails.push(email);
      }
    }
    if (emails.length > 0) {
      vendorEmailMap.set(vendorKey, emails);
    }
  }

  console.log(`Resolved contact emails for ${vendorEmailMap.size} vendors`);
  return vendorEmailMap;
}


/** ========== HELPERS ========== **/

function normalizeStatus_(s) {
  const v = String(s || '').trim().toLowerCase();
  if (!v) return 'other';
  if (v.includes('live')) return 'live';
  if (v.includes('onboard')) return 'onboarding';
  if (v.includes('pause')) return 'paused';
  if (v.includes('pre')) return 'preonboarding';
  if (v.includes('early')) return 'early talks';
  if (v.includes('top') && v.includes('500')) return 'top 500 remodelers';
  if (v.includes('dead') || (v.includes('no') && v.includes('go')) || v.includes('closed')) return 'dead';
  return 'other';
}

function readBlacklist_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const set = new Set();
  if (!sh) return set;
  const last = sh.getLastRow();
  if (last < 2) return set;
  const header = String(sh.getRange(1, 5).getValue() || '').trim().toLowerCase();
  if (header !== 'blacklist' && header !== 'vendor blacklist') return set;
  const vals = sh.getRange(2, 5, last - 1, 1).getValues().flat();
  for (const v of vals) {
    const norm = normalizeName_(v);
    if (norm) set.add(norm.toLowerCase());
  }
  return set;
}

function normalizeName_(v) {
  if (v == null) return '';
  let s = String(v).trim();
  s = s.replace(/^\[\s*\d+\s*\]\s*/, '');
  return s.trim();
}

/** De-dupe by lowercased name, keeping the first occurrence */
function dedupeKeepFirst_(rows) {
  const seen = new Set();
  return rows.filter(r => {
    const key = r.name.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function eq_(a, b) {
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Required sheet "${name}" not found`);
  return sh;
}
