/***** CONFIG *****/
const SHEET_BUYERS_L1M       = 'Buyers L1M';
const SHEET_BUYERS_L6M       = 'Buyers L6M';
const SHEET_AFFILIATES_L1M   = 'Affiliates L1M';
const SHEET_AFFILIATES_L6M   = 'Affiliates L6M';
const SHEET_MON_BUYERS       = 'buyers monday.com';
const SHEET_MON_AFFILIATES   = 'affiliates monday.com';

const SHEET_OUT              = 'List';
const SHEET_SETTINGS         = 'Settings';

// Header candidates to auto-detect vendor and TTL
const VENDOR_HEADER_CANDIDATES = ['Buyer', 'Affiliate', 'Publisher', 'Name', 'Vendor'];
const TTL_HEADER_CANDIDATES    = ['TTL , USD', 'TTL, USD', 'TTL USD'];

// monday.com special columns (0-based indexes)
const ALIAS_COL_INDEX = 15;           // Column P (1-based 16) -> 0-based 15
const BUYERS_STATUS_INDEX = 18;       // Column S (1-based 19) -> 0-based 18
const AFFILIATES_STATUS_INDEX = 18;   // Column S (1-based 19) -> 0-based 18
const BUYERS_NOTES_INDEX = 3;         // Column D (1-based 4) -> 0-based 3
const AFFILIATES_NOTES_INDEX = 3;     // Column D (1-based 4) -> 0-based 3

// Status priority used when TTL = 0
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
 * NOTE: Menu is defined in BattleStation.gs onOpen() function
 * to avoid collision with multiple onOpen() declarations
 */

/**
 * Build List with Gmail links and monday.com notes
 * LABEL-AGNOSTIC: Uses configured labels from Settings sheet.
 *
 * Generates columns: Vendor | TTL, USD | Source | Status | Notes | Gmail Link | no snoozing | processed?
 *
 * HOT ZONE: Vendors with recent emails (configured live query) at top
 * NORMAL ZONE: All other vendors sorted by TTL, type, alpha
 */
function buildListWithGmailAndNotes() {
  const ss = SpreadsheetApp.getActive();

  console.log('=== BUILD LIST WITH GMAIL & NOTES START (Label-Agnostic) ===');

  // Get all vendors using existing buildVendorList logic
  const shBuyL1 = mustGetSheet_(ss, SHEET_BUYERS_L1M);
  const shBuyL6 = mustGetSheet_(ss, SHEET_BUYERS_L6M);
  const shAffL1 = mustGetSheet_(ss, SHEET_AFFILIATES_L1M);
  const shAffL6 = mustGetSheet_(ss, SHEET_AFFILIATES_L6M);
  const shMonB  = ss.getSheetByName(SHEET_MON_BUYERS);
  const shMonA  = ss.getSheetByName(SHEET_MON_AFFILIATES);

  const blacklist = readBlacklist_(ss);

  // Read all vendors
  const buyersL1M = readMetricSheet_(shBuyL1, 'Buyer', 'Buyers L1M', blacklist);
  const buyersL1MSet = new Set(buyersL1M.map(r => r.name.toLowerCase()));

  const buyersL6MAll = readMetricSheet_(shBuyL6, 'Buyer', 'Buyers L6M', blacklist);
  const buyersL6M = buyersL6MAll.filter(r => !buyersL1MSet.has(r.name.toLowerCase()));

  const buyersExisting = new Set([...buyersL1MSet, ...buyersL6M.map(r => r.name.toLowerCase())]);
  const buyersMon = shMonB ? readMondaySheet_(shMonB, 'Buyer', 'buyers monday.com', blacklist, buyersExisting) : [];

  const affL1M = readMetricSheet_(shAffL1, 'Affiliate', 'Affiliates L1M', blacklist);
  const affL1MSet = new Set(affL1M.map(r => r.name.toLowerCase()));

  const affL6MAll = readMetricSheet_(shAffL6, 'Affiliate', 'Affiliates L6M', blacklist);
  const affL6M = affL6MAll.filter(r => !affL1MSet.has(r.name.toLowerCase()));

  const affExisting = new Set([...affL1MSet, ...affL6M.map(r => r.name.toLowerCase())]);
  const affMon = shMonA ? readMondaySheet_(shMonA, 'Affiliate', 'affiliates monday.com', blacklist, affExisting) : [];

  console.log('Data loaded:', {
    buyersL1M: buyersL1M.length,
    buyersL6M: buyersL6M.length,
    buyersMon: buyersMon.length,
    affL1M: affL1M.length,
    affL6M: affL6M.length,
    affMon: affMon.length
  });

  // Split by TTL
  const gt0 = [];
  const z_buyL6 = [], z_affL6 = [], z_buyL1 = [], z_affL1 = [], z_mon = [];

  const pushByTtl = (arr, zeroTarget) => {
    for (const r of arr) {
      if ((r.ttl || 0) > 0) gt0.push(r);
      else zeroTarget.push(r);
    }
  };

  pushByTtl(buyersL6M, z_buyL6);
  pushByTtl(affL6M, z_affL6);
  pushByTtl(buyersL1M, z_buyL1);
  pushByTtl(affL1M, z_affL1);

  for (const r of buyersMon) ((r.ttl || 0) > 0 ? gt0 : z_mon).push(r);
  for (const r of affMon) ((r.ttl || 0) > 0 ? gt0 : z_mon).push(r);

  // Sort >0 TTL by TTL desc, then type (buyers first), then alpha
  gt0.sort((a, b) => {
    const ttlDiff = (b.ttl || 0) - (a.ttl || 0);
    if (ttlDiff !== 0) return ttlDiff;
    const rankA = (a.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    const rankB = (b.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    if (rankA !== rankB) return rankA - rankB;
    return String(a.name).localeCompare(String(b.name));
  });

  // Sort zero-TTL groups alphabetically
  const alpha = (a, b) => String(a.name).localeCompare(String(b.name));
  z_buyL6.sort(alpha);
  z_affL6.sort(alpha);
  z_buyL1.sort(alpha);
  z_affL1.sort(alpha);

  // Sort monday.com zero-TTL by status rank, then type, then alpha
  z_mon.sort((a, b) => {
    const sA = STATUS_RANK[String(a.status || '').toLowerCase()] ?? STATUS_RANK['other'];
    const sB = STATUS_RANK[String(b.status || '').toLowerCase()] ?? STATUS_RANK['other'];
    if (sA !== sB) return sA - sB;
    const rankA = (a.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    const rankB = (b.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    if (rankA !== rankB) return rankA - rankB;
    return String(a.name).localeCompare(String(b.name));
  });

  // Assemble all vendors
  const all = [...gt0, ...z_buyL6, ...z_affL6, ...z_buyL1, ...z_affL1, ...z_mon];

  console.log('Total before status/notes lookup:', all.length);

  // Build status and notes maps from monday.com sheets
  const statusMaps = buildStatusMaps_(shMonB, shMonA);
  const notesMaps = buildNotesMaps_(shMonB, shMonA);

  console.log('Maps built:', {
    buyersStatus: statusMaps.buyers.size,
    affiliatesStatus: statusMaps.affiliates.size,
    buyersNotes: notesMaps.buyers.size,
    affiliatesNotes: notesMaps.affiliates.size
  });

  // Lookup status and notes for each vendor
  for (const r of all) {
    r.status = lookupStatus_(r.name, r.type, statusMaps);
    r.notes = lookupNotes_(r.name, r.type, notesMaps);
  }

  // HOT ZONE: Detect vendors with emails (using configured labels)
  console.log('Detecting hot vendors...');
  const hotVendorSet = getHotVendorsFromGmail_(all);
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

  // Final list: Hot zone at top, then normal zone
  const finalList = [...hotZone, ...normalZone];

  console.log('Total vendors for output:', finalList.length);

  // Write to List sheet
  const shOut = ensureSheet_(ss, SHEET_OUT);

  // Clear all content and formatting
  const lastRow = shOut.getMaxRows();
  const lastCol = shOut.getMaxColumns();
  if (lastRow > 0 && lastCol > 0) {
    shOut.getRange(1, 1, lastRow, lastCol).clear();
  }

  // Write headers
  shOut.getRange(1, 1, 1, 8).setValues([[
    'Vendor', 'TTL , USD', 'Source', 'Status', 'Notes',
    'Gmail Link', 'no snoozing', 'processed?'
  ]]).setFontWeight('bold').setBackground('#e8f0fe');

  if (finalList.length > 0) {
    // Write data (columns A-E)
    const data = finalList.map(r => [
      r.name,
      r.ttl || 0,
      r.source,
      r.status || '',
      r.notes || ''
    ]);

    shOut.getRange(2, 1, data.length, 5).setValues(data);
    shOut.getRange(2, 2, data.length, 1).setNumberFormat('$#,##0.00;($#,##0.00)');

    // Read Gmail sublabel mappings from Settings sheet
    const gmailSublabelMap = readGmailSublabelMap_(ss);

    // LABEL-AGNOSTIC: Build Gmail links using configured labels
    const gmailData = finalList.map(v => {
      const vendorKey = v.name.toLowerCase();
      const vendorSlug = gmailSublabelMap.has(vendorKey)
        ? gmailSublabelMap.get(vendorKey)
        : null;

      // Use the label config to build queries
      const queries = buildVendorEmailQuery_(v.name, vendorSlug);
      const gmailAll = buildGmailSearchUrl_(queries.allQuery);
      const gmailNoSnooze = buildGmailSearchUrl_(queries.noSnoozeQuery);

      return [gmailAll, gmailNoSnooze, false];
    });

    shOut.getRange(2, 6, gmailData.length, 3).setValues(gmailData);
  }

  // Auto-resize columns
  shOut.autoResizeColumns(1, 5);
  shOut.setColumnWidth(6, 100);
  shOut.setColumnWidth(7, 100);
  shOut.setColumnWidth(8, 100);

  console.log('=== BUILD LIST WITH GMAIL & NOTES END ===');

  const counts = {
    'HOT (recent emails)': hotZone.length,
    'Total vendors': finalList.length,
    'With status': finalList.filter(v => v.status).length,
    'With notes': finalList.filter(v => v.notes).length
  };

  console.log('Final counts:', counts);

  ss.toast(
    Object.entries(counts).map(([k,v]) => `${k}: ${v}`).join(' | '),
    'List Built',
    8
  );
}


/** ========== HOT ZONE DETECTION (Label-Agnostic) ========== **/

/**
 * Search Gmail for threads that indicate "hot" vendors.
 * LABEL-AGNOSTIC: Uses configured live_email_query and vendor_label_prefix.
 *
 * A vendor is "hot" if they have emails matching the configured live query.
 *
 * Detection methods (in priority order):
 * 1. Configured vendor labels (if vendor_label_style is not "none")
 * 2. Exact vendor name match in subject/sender/recipient
 */
function getHotVendorsFromGmail_(allVendors) {
  const hotSet = new Set();
  const cfg = getLabelConfig_();

  try {
    // Build vendor name lookup for fast matching
    const vendorMap = new Map();
    const vendorNames = allVendors.map(v => {
      const nameLower = v.name.toLowerCase();
      vendorMap.set(nameLower, v.name);
      return {
        name: v.name,
        nameLower: nameLower
      };
    });

    let labelMatches = 0;
    let exactMatches = 0;

    // SEARCH 1: Emails matching the configured live query
    const liveQuery = cfg.live_email_query || 'label:inbox';
    console.log(`Searching for emails with: ${liveQuery}`);
    const recentThreads = GmailApp.search(liveQuery, 0, 200);
    console.log(`Found ${recentThreads.length} threads with live query`);

    // SEARCH 2: All inbox emails (as fallback)
    console.log('Searching inbox for any vendor emails...');
    const inboxThreads = GmailApp.search('in:inbox', 0, 200);
    console.log(`Found ${inboxThreads.length} inbox threads`);

    // Combine both searches (dedupe by thread ID)
    const processedThreadIds = new Set();
    const allThreads = [...recentThreads, ...inboxThreads];

    // Determine vendor label prefix for matching
    const vendorLabelPrefix = cfg.vendor_label_prefix || '';
    const vendorLabelStyle = cfg.vendor_label_style || 'none';

    for (const thread of allThreads) {
      try {
        const threadId = thread.getId();

        // Skip if we've already processed this thread
        if (processedThreadIds.has(threadId)) continue;
        processedThreadIds.add(threadId);

        let matched = false;

        // METHOD 1: Check for vendor labels (if configured)
        if (vendorLabelStyle !== 'none' && vendorLabelPrefix) {
          const labels = thread.getLabels();
          for (const label of labels) {
            const labelName = label.getName();

            if (vendorLabelStyle === 'sublabel' && labelName.startsWith(vendorLabelPrefix)) {
              // Sublabel style: prefix/VendorName
              const vendorNameFromLabel = labelName.substring(vendorLabelPrefix.length).toLowerCase();

              if (vendorMap.has(vendorNameFromLabel)) {
                hotSet.add(vendorNameFromLabel);
                labelMatches++;
                matched = true;
                break;
              }

              // Partial match
              for (const vendor of vendorNames) {
                if (vendorNameFromLabel.includes(vendor.nameLower) ||
                    vendor.nameLower.includes(vendorNameFromLabel)) {
                  hotSet.add(vendor.nameLower);
                  labelMatches++;
                  matched = true;
                  break;
                }
              }
              if (matched) break;

            } else if (vendorLabelStyle === 'flat') {
              // Flat style: prefix-vendor-slug
              const labelLower = labelName.toLowerCase();
              if (labelLower.startsWith(vendorLabelPrefix.toLowerCase())) {
                const slugFromLabel = labelLower.substring(vendorLabelPrefix.length);

                for (const vendor of vendorNames) {
                  const vendorSlug = vendor.nameLower
                    .replace(/[^a-z0-9-]+/g, '-')
                    .replace(/^-+|-+$/g, '');
                  if (slugFromLabel === vendorSlug) {
                    hotSet.add(vendor.nameLower);
                    labelMatches++;
                    matched = true;
                    break;
                  }
                }
                if (matched) break;
              }
            }
          }
        }

        if (matched) continue;

        // METHOD 2: Exact name match in subject/sender/recipient
        const subject = thread.getFirstMessageSubject().toLowerCase();
        const messages = thread.getMessages();

        let emailText = subject;
        if (messages.length > 0) {
          const firstMsg = messages[0];
          emailText += ' ' + firstMsg.getFrom().toLowerCase();
          emailText += ' ' + firstMsg.getTo().toLowerCase();
        }

        for (const vendor of vendorNames) {
          if (emailText.includes(vendor.nameLower)) {
            hotSet.add(vendor.nameLower);
            exactMatches++;
            matched = true;
            break;
          }
        }
      } catch (e) {
        console.log(`Error processing thread: ${e.message}`);
      }
    }

    console.log(`Hot vendor detection: ${labelMatches} label matches, ${exactMatches} exact matches`);
    console.log(`Total unique threads processed: ${processedThreadIds.size}`);

  } catch (e) {
    console.log(`Error searching Gmail: ${e.message}`);
  }

  return hotSet;
}


/** ========== STATUS & NOTES LOOKUP ========== **/

/**
 * Build status lookup maps from monday.com sheets
 * Returns { buyers: Map, affiliates: Map }
 */
function buildStatusMaps_(shMonB, shMonA) {
  const buyersMap = new Map();
  const affiliatesMap = new Map();

  if (shMonB) {
    const buyersData = shMonB.getDataRange().getValues();
    for (let i = 1; i < buyersData.length; i++) {
      const row = buyersData[i];
      const vendor = normalizeName_(row[0]);
      const status = BUYERS_STATUS_INDEX < row.length ? String(row[BUYERS_STATUS_INDEX] || '').trim() : '';
      if (vendor) {
        buyersMap.set(vendor.toLowerCase(), status);
      }
    }
  }

  if (shMonA) {
    const affiliatesData = shMonA.getDataRange().getValues();
    for (let i = 1; i < affiliatesData.length; i++) {
      const row = affiliatesData[i];
      const vendor = normalizeName_(row[0]);
      const status = AFFILIATES_STATUS_INDEX < row.length ? String(row[AFFILIATES_STATUS_INDEX] || '').trim() : '';
      if (vendor) {
        affiliatesMap.set(vendor.toLowerCase(), status);
      }
    }
  }

  return { buyers: buyersMap, affiliates: affiliatesMap };
}

/**
 * Build notes lookup maps from monday.com sheets
 * Returns { buyers: Map, affiliates: Map }
 */
function buildNotesMaps_(shMonB, shMonA) {
  const buyersMap = new Map();
  const affiliatesMap = new Map();

  if (shMonB) {
    const buyersData = shMonB.getDataRange().getValues();
    for (let i = 1; i < buyersData.length; i++) {
      const row = buyersData[i];
      const vendor = normalizeName_(row[0]);
      const notes = BUYERS_NOTES_INDEX < row.length ? String(row[BUYERS_NOTES_INDEX] || '').trim() : '';
      if (vendor) {
        buyersMap.set(vendor.toLowerCase(), notes);
      }
    }
  }

  if (shMonA) {
    const affiliatesData = shMonA.getDataRange().getValues();
    for (let i = 1; i < affiliatesData.length; i++) {
      const row = affiliatesData[i];
      const vendor = normalizeName_(row[0]);
      const notes = AFFILIATES_NOTES_INDEX < row.length ? String(row[AFFILIATES_NOTES_INDEX] || '').trim() : '';
      if (vendor) {
        affiliatesMap.set(vendor.toLowerCase(), notes);
      }
    }
  }

  return { buyers: buyersMap, affiliates: affiliatesMap };
}

/**
 * Lookup status for a vendor from the status maps
 */
function lookupStatus_(name, type, statusMaps) {
  const key = name.toLowerCase();
  const isBuyer = (type || '').toLowerCase().startsWith('buyer');

  if (isBuyer && statusMaps.buyers.has(key)) {
    return normalizeStatus_(statusMaps.buyers.get(key));
  }
  if (!isBuyer && statusMaps.affiliates.has(key)) {
    return normalizeStatus_(statusMaps.affiliates.get(key));
  }
  if (statusMaps.buyers.has(key)) {
    return normalizeStatus_(statusMaps.buyers.get(key));
  }
  if (statusMaps.affiliates.has(key)) {
    return normalizeStatus_(statusMaps.affiliates.get(key));
  }
  return '';
}

/**
 * Lookup notes for a vendor from the notes maps
 */
function lookupNotes_(name, type, notesMaps) {
  const key = name.toLowerCase();
  const isBuyer = (type || '').toLowerCase().startsWith('buyer');

  if (isBuyer && notesMaps.buyers.has(key)) {
    return notesMaps.buyers.get(key);
  }
  if (!isBuyer && notesMaps.affiliates.has(key)) {
    return notesMaps.affiliates.get(key);
  }
  if (notesMaps.buyers.has(key)) {
    return notesMaps.buyers.get(key);
  }
  if (notesMaps.affiliates.has(key)) {
    return notesMaps.affiliates.get(key);
  }
  return '';
}


/** ========== GMAIL SUBLABEL MAPPING ========== **/

/**
 * Read Gmail sublabel mappings from Settings sheet
 * Returns a Map of lowercase vendor name -> gmail sublabel
 */
function readGmailSublabelMap_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const map = new Map();

  if (!sh) return map;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return map;

  // Assuming Column A = Vendor Name, Column B = Gmail Sublabel
  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();

  for (const row of data) {
    const vendorName = String(row[0] || '').trim();
    const sublabel = String(row[1] || '').trim();

    if (vendorName && sublabel) {
      map.set(vendorName.toLowerCase(), sublabel);
    }
  }

  console.log(`Loaded ${map.size} Gmail sublabel mappings from Settings`);
  return map;
}


/** ========== READERS & HELPERS ========== **/

/**
 * Read a metric sheet (L1M/L6M) and return [{name, ttl, type, source}, ...]
 */
function readMetricSheet_(sh, type, sourceLabel, blacklist) {
  const { columns, rows } = readObjectsFromSheet_(sh);
  console.log(`[${sourceLabel}] Total rows:`, rows.length);

  if (!rows.length) return [];
  const vHeader = firstExistingHeader_(columns, VENDOR_HEADER_CANDIDATES, sh.getName(), 'vendor');
  const tHeader = firstExistingHeader_(columns, TTL_HEADER_CANDIDATES, sh.getName(), 'TTL , USD');

  const out = [];
  for (const r of rows) {
    const name = normalizeName_(r[vHeader]);
    if (!name) continue;

    const key = name.toLowerCase();
    if (blacklist.has(key)) {
      recordSkipReason_({ name, source: sourceLabel, reason: 'blacklist' });
      continue;
    }

    out.push({
      name,
      ttl: toNumber_(r[tHeader]),
      type,
      source: sourceLabel
    });
  }
  return dedupeKeepMaxTTL_(out);
}

/**
 * Read a monday.com sheet
 */
function readMondaySheet_(sh, type, sourceLabel, blacklist, existingSet) {
  console.log(`\n[${sourceLabel}] START READ`);

  const allValues = sh.getDataRange().getValues();
  if (allValues.length < 2) return [];

  const headers = allValues[0].map(h => String(h || '').trim());
  console.log(`[${sourceLabel}] Total rows:`, allValues.length - 1);

  const nameIdx = 0;

  const ttlIdx = headers.findIndex(h =>
    eq_(h, 'TTL , USD') || eq_(h, 'TTL, USD') || eq_(h, 'TTL USD')
  );
  const typeIdx = headers.findIndex(h => eq_(h, 'Type'));

  const statusIdx = BUYERS_STATUS_INDEX;
  const notesIdx = BUYERS_NOTES_INDEX;

  const out = [];
  let skippedBlacklist = 0;
  let skippedExisting = 0;
  let skippedAlias = 0;
  let skippedNoName = 0;

  for (let i = 1; i < allValues.length; i++) {
    const row = allValues[i];

    const rawName = row[nameIdx] || '';
    const name = normalizeName_(rawName);

    if (!name) {
      skippedNoName++;
      continue;
    }

    const key = name.toLowerCase();

    if (blacklist.has(key)) {
      skippedBlacklist++;
      continue;
    }

    if (existingSet.has(key)) {
      skippedExisting++;
      continue;
    }

    const aliasRaw = (ALIAS_COL_INDEX < row.length) ? (row[ALIAS_COL_INDEX] || '') : '';
    const aliases = String(aliasRaw).split(',')
      .map(s => normalizeName_(s).toLowerCase())
      .filter(Boolean);

    if (aliases.some(a => existingSet.has(a) || blacklist.has(a))) {
      skippedAlias++;
      recordSkipReason_({ name, source: sourceLabel, reason: 'alias-duplicate' });
      continue;
    }

    const ttl = (ttlIdx !== -1 && ttlIdx < row.length) ? toNumber_(row[ttlIdx]) : 0;
    const explicitType = (typeIdx !== -1 && typeIdx < row.length) ? String(row[typeIdx] || '').trim() : '';
    const finalType = explicitType || type;
    const statusRaw = (statusIdx != null && statusIdx < row.length) ? String(row[statusIdx] || '').trim() : '';
    const status = normalizeStatus_(statusRaw);
    const notes = (notesIdx != null && notesIdx < row.length) ? String(row[notesIdx] || '').trim() : '';

    out.push({ name, ttl, type: finalType, source: sourceLabel, status, notes });
    existingSet.add(key);
  }

  console.log(`[${sourceLabel}] Skip summary:`, {
    noName: skippedNoName,
    blacklist: skippedBlacklist,
    existing: skippedExisting,
    alias: skippedAlias,
    added: out.length
  });
  console.log(`[${sourceLabel}] END READ\n`);

  return dedupeKeepMaxTTL_(out);
}

/** Normalize buyer or affiliate status to our buckets */
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

/** Reads Settings column E "blacklist" and returns a Set of normalized, lowercased names */
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

/** Data readers */
function readObjectsFromSheet_(sh) {
  const values = sh.getDataRange().getValues();
  if (!values.length) return { columns: [], rows: [] };
  const headers = values[0].map(h => String(h || '').trim());
  const rows = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    return obj;
  });
  return { columns: headers, rows };
}

/** De-dupe by lowercased name keeping max TTL */
function dedupeKeepMaxTTL_(rows) {
  const map = new Map();
  for (const r of rows) {
    const key = r.name.toLowerCase();
    const prev = map.get(key);
    if (!prev || (r.ttl || 0) > (prev.ttl || 0)) {
      map.set(key, r);
    }
  }
  return Array.from(map.values());
}

/** Normalize display name by removing leading "[123]" token if present */
function normalizeName_(v) {
  if (v == null) return '';
  let s = String(v).trim();
  s = s.replace(/^\[\s*\d+\s*\]\s*/, '');
  return s.trim();
}

/** Robust currency parsing */
function toNumber_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  let s = String(v).trim();
  let neg = false;
  if (s.startsWith('(') && s.endsWith(')')) {
    neg = true; s = s.slice(1, -1);
  }
  s = s.replace(/[$,\s]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : (neg ? -n : n);
}

/** Case-insensitive string equality */
function eq_(a, b) {
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

/** Find the first header that exists from a list of candidates */
function firstExistingHeader_(columns, candidates, sheetName, purpose) {
  for (const c of candidates) {
    if (columns.some(col => eq_(col, c))) {
      return c;
    }
  }
  console.log(`Warning: No matching header found for ${purpose} in ${sheetName}. Using first candidate: ${candidates[0]}`);
  return candidates[0];
}

/** Find header index (optional, returns null if not found) */
function findHeaderOptionalIndex_(columns, header) {
  const idx = columns.findIndex(col => eq_(col, header));
  return idx === -1 ? null : idx;
}

/** Ensure a sheet exists, create if not */
function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  return sh;
}

/** Get a sheet, throw if not found */
function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error(`Required sheet "${name}" not found`);
  }
  return sh;
}

/** Record skip reasons for debugging */
function recordSkipReason_(obj) {
  SKIP_REASONS.push(obj);
}

/** Get skip reasons (for debugging) */
function getSkipReasons() {
  return SKIP_REASONS;
}
