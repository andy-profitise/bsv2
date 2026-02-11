/************************************************************
 * LABEL CONFIG - Label-Agnostic Settings System
 *
 * Instead of hardcoded Gmail labels, all label references
 * are read from the Settings sheet. Users configure which
 * labels define their "live" emails, vendor matching, etc.
 *
 * Settings Sheet Layout (Columns G-I):
 *   G: Setting Key
 *   H: Setting Value
 *   I: Description
 *
 * This makes the Battle Station work with ANY Gmail label
 * structure, not just the original 00.received/zzzVendors.
 ************************************************************/

const LABEL_CFG_DEFAULTS = {
  // The Gmail search query for "live" emails (what shows up as active)
  live_email_query: 'label:inbox',

  // Prefix for vendor-specific Gmail labels (e.g., "zzzVendors/" or "vendors/" or "")
  vendor_label_prefix: '',

  // How vendor labels are formatted:
  //   "sublabel"  = prefix/Vendor Name  (e.g., zzzVendors/Acme Corp)
  //   "flat"      = prefix-vendor-slug  (e.g., zzzvendors-acme-corp)
  //   "none"      = no vendor labels (match by name/email in thread content)
  vendor_label_style: 'none',

  // Gmail query fragment for snoozed emails
  snoozed_query: 'is:snoozed',

  // Gmail query fragment to exclude certain emails (e.g., "-label:archived")
  exclude_query: '',

  // Label name for priority emails
  priority_label: '',

  // Label name for "waiting on customer" emails
  waiting_customer_label: '',

  // Label name for "waiting on me" emails
  waiting_me_label: '',

  // Label name for "waiting on external party" emails
  waiting_external_label: '',

  // Label name for accounting/invoice emails
  accounting_label: '',

  // Google Sheet ID for email logging (blank = disabled)
  email_log_spreadsheet_id: '',

  // Business hours before priority+waiting email is overdue
  overdue_business_hours: '16',

  // Days before "waiting on external" email is overdue
  overdue_external_days: '7',

  // How to search for vendor emails when no vendor labels exist:
  //   "name"     = search by vendor name in subject/body
  //   "contacts" = search by contact email addresses
  //   "both"     = search by both name and contacts
  vendor_search_mode: 'both'
};

// Column positions in Settings sheet for label config (1-based)
const LABEL_CFG_KEY_COL = 7;    // Column G
const LABEL_CFG_VALUE_COL = 8;  // Column H
const LABEL_CFG_DESC_COL = 9;   // Column I

// Cache for label config (within a single execution)
let _labelConfigCache = null;

/**
 * Get the full label configuration from Settings sheet.
 * Falls back to defaults for any missing keys.
 * Results are cached for the duration of the script execution.
 *
 * @returns {Object} Configuration object with all label settings
 */
function getLabelConfig_() {
  if (_labelConfigCache) return _labelConfigCache;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Settings');
  const config = Object.assign({}, LABEL_CFG_DEFAULTS);

  if (!sh) {
    _labelConfigCache = config;
    return config;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    _labelConfigCache = config;
    return config;
  }

  // Read columns G-H for key-value pairs
  const data = sh.getRange(2, LABEL_CFG_KEY_COL, lastRow - 1, 2).getValues();

  for (const row of data) {
    const key = String(row[0] || '').trim().toLowerCase();
    const value = String(row[1] || '').trim();

    if (key && value !== '' && LABEL_CFG_DEFAULTS.hasOwnProperty(key)) {
      config[key] = value;
    }
  }

  _labelConfigCache = config;
  return config;
}

/**
 * Clear the label config cache (call when settings change)
 */
function clearLabelConfigCache_() {
  _labelConfigCache = null;
}

/**
 * Build a Gmail search query for a vendor's live emails.
 *
 * Depending on config, this could be:
 *   - "label:inbox label:zzzVendors/Acme Corp"  (sublabel style)
 *   - "label:inbox label:zzzvendors-acme-corp"   (flat style)
 *   - "label:inbox Acme Corp"                    (name match, no vendor labels)
 *
 * @param {string} vendorName - The vendor name
 * @param {string} [vendorSlug] - Optional pre-computed slug for flat labels
 * @param {Array} [contactEmails] - Optional contact email addresses for search
 * @returns {Object} { allQuery, noSnoozeQuery } - Gmail search queries
 */
function buildVendorEmailQuery_(vendorName, vendorSlug, contactEmails) {
  const cfg = getLabelConfig_();

  let vendorFilter = '';

  if (cfg.vendor_label_style === 'sublabel' && cfg.vendor_label_prefix) {
    // Sublabel: label:prefix/VendorName
    vendorFilter = `label:${cfg.vendor_label_prefix}${vendorName}`;
  } else if (cfg.vendor_label_style === 'flat' && cfg.vendor_label_prefix) {
    // Flat: label:prefix-vendor-slug
    const slug = vendorSlug || vendorName.toLowerCase()
      .replace(/[^a-z0-9-]+/g, '-')
      .replace(/^-+|-+$/g, '');
    vendorFilter = `label:${cfg.vendor_label_prefix}${slug}`;
  } else {
    // No vendor labels: search by name and/or contact emails
    const parts = [];

    if (cfg.vendor_search_mode === 'name' || cfg.vendor_search_mode === 'both') {
      // Use quotes for exact name match in search
      parts.push(`"${vendorName}"`);
    }

    if ((cfg.vendor_search_mode === 'contacts' || cfg.vendor_search_mode === 'both') &&
        contactEmails && contactEmails.length > 0) {
      // Search by contact emails: from:email1 OR from:email2
      const emailFilters = contactEmails.map(e => `from:${e} OR to:${e}`).join(' OR ');
      if (emailFilters) {
        parts.push(`(${emailFilters})`);
      }
    }

    if (parts.length > 1) {
      vendorFilter = `{${parts.join(' ')}}`;
    } else if (parts.length === 1) {
      vendorFilter = parts[0];
    }
  }

  // Build the full query
  const liveQuery = cfg.live_email_query || 'label:inbox';
  const excludeQuery = cfg.exclude_query || '';

  const allQuery = [liveQuery, vendorFilter, excludeQuery].filter(Boolean).join(' ');
  const noSnoozeQuery = [allQuery, `-${cfg.snoozed_query || 'is:snoozed'}`].filter(Boolean).join(' ');

  return { allQuery, noSnoozeQuery };
}

/**
 * Build a Gmail search link URL from a query
 *
 * @param {string} query - Gmail search query
 * @returns {string} Gmail URL
 */
function buildGmailSearchUrl_(query) {
  return `https://mail.google.com/mail/u/0/#search/${encodeURIComponent(query)}`;
}

/**
 * Check if an email has a specific configured label.
 *
 * @param {string} emailLabels - Comma-separated label string from email
 * @param {string} configKey - Key in label config (e.g., 'priority_label')
 * @returns {boolean}
 */
function hasConfiguredLabel_(emailLabels, configKey) {
  const cfg = getLabelConfig_();
  const labelName = cfg[configKey];

  if (!labelName || !emailLabels) return false;

  return emailLabels.split(',').map(l => l.trim()).includes(labelName);
}

/**
 * Check if an email is overdue based on configured labels.
 * Uses the configured label names instead of hardcoded ones.
 *
 * @param {Object} email - Email object with .labels and .date
 * @returns {boolean}
 */
function isEmailOverdueConfigurable_(email) {
  if (!email || !email.labels) return false;

  const cfg = getLabelConfig_();
  const emailDate = parseEmailDate_(email.date);
  if (!emailDate) return false;

  // Check for external waiting - overdue after N calendar days
  if (cfg.waiting_external_label && email.labels.includes(cfg.waiting_external_label)) {
    const now = new Date();
    const daysDiff = (now - emailDate) / (1000 * 60 * 60 * 24);
    const maxDays = parseInt(cfg.overdue_external_days) || 7;
    if (daysDiff > maxDays) return true;
  }

  // Check for priority + waiting/customer or waiting/me - N business hours
  if (cfg.priority_label && email.labels.includes(cfg.priority_label)) {
    const isWaiting = (cfg.waiting_customer_label && email.labels.includes(cfg.waiting_customer_label)) ||
                      (cfg.waiting_me_label && email.labels.includes(cfg.waiting_me_label));
    if (isWaiting) {
      const businessHours = getBusinessHoursElapsed_(emailDate);
      const maxHours = parseInt(cfg.overdue_business_hours) || 16;
      if (businessHours > maxHours) return true;
    }
  }

  return false;
}

/**
 * Determine email status color based on configured labels.
 * Returns a color category for the email.
 *
 * @param {Object} email - Email object with .labels
 * @returns {string} Status category: 'snoozed', 'overdue', 'priority_waiting',
 *                   'waiting_external', 'accounting', 'waiting_customer', 'normal'
 */
function getEmailStatusCategory_(email) {
  if (!email) return 'normal';

  const cfg = getLabelConfig_();

  if (email.isSnoozed) return 'snoozed';

  if (isEmailOverdueConfigurable_(email)) return 'overdue';

  // Priority label
  const hasPriority = cfg.priority_label && email.labels && email.labels.includes(cfg.priority_label);

  if (hasPriority) {
    if (cfg.waiting_external_label && email.labels.includes(cfg.waiting_external_label)) {
      return 'waiting_external';
    }
    if (cfg.accounting_label && email.labels.includes(cfg.accounting_label)) {
      return 'accounting';
    }
    if (cfg.waiting_customer_label && email.labels.includes(cfg.waiting_customer_label)) {
      return 'waiting_customer';
    }
    return 'priority_waiting';
  }

  // Non-priority but labeled
  if (cfg.waiting_external_label && email.labels && email.labels.includes(cfg.waiting_external_label)) {
    return 'waiting_external';
  }
  if (cfg.accounting_label && email.labels && email.labels.includes(cfg.accounting_label)) {
    return 'accounting';
  }

  return 'normal';
}

/**
 * Set up the label configuration section in the Settings sheet.
 * Creates headers and populates defaults if the section doesn't exist.
 */
function setupLabelConfig() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Settings');

  if (!sh) {
    sh = ss.insertSheet('Settings');
  }

  // Check if label config section already exists
  const lastRow = sh.getLastRow();
  let hasConfig = false;

  if (lastRow >= 2) {
    const keys = sh.getRange(2, LABEL_CFG_KEY_COL, lastRow - 1, 1).getValues().flat();
    hasConfig = keys.some(k => String(k).trim().toLowerCase() === 'live_email_query');
  }

  if (hasConfig) {
    SpreadsheetApp.getUi().alert('Label configuration already exists in Settings sheet.\n\nEdit the values in columns G-H to change settings.');
    return;
  }

  // Write header
  sh.getRange(1, LABEL_CFG_KEY_COL).setValue('Setting Key').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(1, LABEL_CFG_VALUE_COL).setValue('Setting Value').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(1, LABEL_CFG_DESC_COL).setValue('Description').setFontWeight('bold').setBackground('#e8f0fe');

  // Write all config entries with defaults and descriptions
  const entries = [
    ['live_email_query', LABEL_CFG_DEFAULTS.live_email_query, 'Gmail query for live/active emails (e.g., "label:inbox", "label:unread", "label:00.received")'],
    ['vendor_label_prefix', LABEL_CFG_DEFAULTS.vendor_label_prefix, 'Prefix for vendor-specific Gmail labels (e.g., "zzzVendors/", "vendors/", or leave blank)'],
    ['vendor_label_style', LABEL_CFG_DEFAULTS.vendor_label_style, '"sublabel" (prefix/Name), "flat" (prefix-slug), or "none" (search by name/email)'],
    ['vendor_search_mode', LABEL_CFG_DEFAULTS.vendor_search_mode, 'When vendor_label_style=none: "name", "contacts", or "both"'],
    ['snoozed_query', LABEL_CFG_DEFAULTS.snoozed_query, 'Gmail query to identify snoozed emails (e.g., "is:snoozed")'],
    ['exclude_query', LABEL_CFG_DEFAULTS.exclude_query, 'Gmail query to exclude emails (e.g., "-label:archived")'],
    ['priority_label', LABEL_CFG_DEFAULTS.priority_label, 'Label name for priority emails (leave blank if not used)'],
    ['waiting_customer_label', LABEL_CFG_DEFAULTS.waiting_customer_label, 'Label for "waiting on customer" (leave blank if not used)'],
    ['waiting_me_label', LABEL_CFG_DEFAULTS.waiting_me_label, 'Label for "waiting on me" (leave blank if not used)'],
    ['waiting_external_label', LABEL_CFG_DEFAULTS.waiting_external_label, 'Label for "waiting on external party" (leave blank if not used)'],
    ['accounting_label', LABEL_CFG_DEFAULTS.accounting_label, 'Label for accounting/invoice emails (leave blank if not used)'],
    ['email_log_spreadsheet_id', LABEL_CFG_DEFAULTS.email_log_spreadsheet_id, 'Google Sheet ID for email logging (blank = create new)'],
    ['overdue_business_hours', LABEL_CFG_DEFAULTS.overdue_business_hours, 'Business hours before priority+waiting email is overdue'],
    ['overdue_external_days', LABEL_CFG_DEFAULTS.overdue_external_days, 'Days before "waiting on external" email is overdue'],
  ];

  if (entries.length > 0) {
    sh.getRange(2, LABEL_CFG_KEY_COL, entries.length, 3).setValues(entries);
  }

  // Auto-resize description column
  sh.setColumnWidth(LABEL_CFG_KEY_COL, 200);
  sh.setColumnWidth(LABEL_CFG_VALUE_COL, 200);
  sh.setColumnWidth(LABEL_CFG_DESC_COL, 400);

  // Wrap description column
  sh.getRange(2, LABEL_CFG_DESC_COL, entries.length, 1).setWrap(true);

  SpreadsheetApp.getUi().alert(
    'Label configuration created in Settings sheet (columns G-I).\n\n' +
    'Edit the "Setting Value" column (H) to configure your Gmail label structure.\n\n' +
    'Default: Searches label:inbox with name-based vendor matching.'
  );
}
