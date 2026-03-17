/**
 * LinkManagement.gs
 * Micro-Phase 21: Link Management System
 *
 * Allows Directors and Operators to add, edit, remove, and reorder external links
 * displayed in the Quick Links panel on the dashboard.
 *
 * Links Sheet Structure:
 * link_id | title | url | description | icon | category | house | display_order | status | added_by | added_date | updated_by | updated_date
 */

// Fallback for SPREADSHEET_ID if not defined globally
const LINK_SPREADSHEET_ID = typeof SPREADSHEET_ID !== 'undefined' ? SPREADSHEET_ID : '1w71ytbfftinyG2GeAdM6NDlFjpbHCChVH1V49CObmkc';
const LINK_SHEET_PRIMARY_NAME = 'Link_Management';
const LINK_SHEET_FALLBACK_NAME = 'Links';

const LINK_HEADERS_SIMPLE = [
  'Link_ID',
  'Link_Name',
  'Link_URL',
  'Icon',
  'House',
  'Order',
  'Active'
];

const LINK_HEADERS_LEGACY = [
  'link_id',
  'title',
  'url',
  'description',
  'icon',
  'category',
  'house',
  'display_order',
  'status',
  'added_by',
  'added_date',
  'updated_by',
  'updated_date'
];

function normalizeLinkHeader_(header) {
  return String(header || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getLinkHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const normalized = normalizeLinkHeader_(header);
    const keyMap = {
      linkid: 'link_id',
      linkname: 'title',
      title: 'title',
      linkurl: 'url',
      url: 'url',
      icon: 'icon',
      house: 'house',
      order: 'display_order',
      displayorder: 'display_order',
      active: 'active',
      status: 'status',
      description: 'description',
      category: 'category',
      addedby: 'added_by',
      addeddate: 'added_date',
      updatedby: 'updated_by',
      updateddate: 'updated_date'
    };
    const key = keyMap[normalized];
    if (key) {
      map[key] = index;
    }
  });
  return map;
}

function parseActiveValue_(value) {
  if (typeof value === 'boolean') return value;
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return true;
  return ['true', 'yes', 'active', '1'].includes(normalized);
}

function buildLinkRow_(headers, linkData, metadata) {
  const row = new Array(headers.length).fill('');
  const headerMap = getLinkHeaderMap_(headers);
  const active = typeof linkData.active === 'boolean'
    ? linkData.active
    : String(linkData.status || 'Active').toLowerCase() === 'active';
  const status = active ? 'Active' : 'Hidden';

  Object.keys(headerMap).forEach(key => {
    const idx = headerMap[key];
    switch (key) {
      case 'link_id':
        row[idx] = linkData.link_id || '';
        break;
      case 'title':
        row[idx] = linkData.title || '';
        break;
      case 'url':
        row[idx] = linkData.url || '';
        break;
      case 'icon':
        row[idx] = linkData.icon || '';
        break;
      case 'house':
        row[idx] = linkData.house || 'Both';
        break;
      case 'display_order':
        row[idx] = linkData.display_order || '';
        break;
      case 'active':
        row[idx] = active ? true : false;
        break;
      case 'status':
        row[idx] = status;
        break;
      case 'description':
        row[idx] = linkData.description || '';
        break;
      case 'category':
        row[idx] = linkData.category || 'General';
        break;
      case 'added_by':
        row[idx] = metadata.added_by || '';
        break;
      case 'added_date':
        row[idx] = metadata.added_date || '';
        break;
      case 'updated_by':
        row[idx] = metadata.updated_by || '';
        break;
      case 'updated_date':
        row[idx] = metadata.updated_date || '';
        break;
    }
  });
  return row;
}

/**
 * Get the Links sheet, creating it if it doesn't exist
 * Also handles migration from old 3-column structure to new 12-column structure
 */
function getLinksSheet() {
  try {
    const ss = SpreadsheetApp.openById(LINK_SPREADSHEET_ID);
    if (!ss) {
      throw new Error('Could not open spreadsheet with ID: ' + LINK_SPREADSHEET_ID);
    }
    let sheet = ss.getSheetByName(LINK_SHEET_PRIMARY_NAME) || ss.getSheetByName(LINK_SHEET_FALLBACK_NAME);
    const expectedHeaders = LINK_HEADERS_SIMPLE;

    if (!sheet) {
      sheet = ss.insertSheet(LINK_SHEET_PRIMARY_NAME);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      addDefaultLinks(sheet);
      applyLinkColumnValidation_(sheet, expectedHeaders);
    } else {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) {
        sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
        sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold');
        sheet.setFrozenRows(1);
        addDefaultLinks(sheet);
        applyLinkColumnValidation_(sheet, expectedHeaders);
      } else {
        const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const normalizedHeaders = currentHeaders.map(normalizeLinkHeader_);
        const hasLinkId = normalizedHeaders.includes('linkid');
        const hasHouse = normalizedHeaders.includes('house');
        const hasActive = normalizedHeaders.includes('active');

        if (!hasLinkId && (normalizedHeaders.includes('label') || normalizedHeaders.includes('title') || currentHeaders.length < 6)) {
          console.log('Migrating Links sheet from old structure...');
          migrateLinksSheet(sheet, currentHeaders, expectedHeaders);
        } else if (!hasHouse) {
          const insertIndex = normalizedHeaders.indexOf('order');
          const houseCol = insertIndex === -1 ? currentHeaders.length + 1 : insertIndex + 1;
          sheet.insertColumnBefore(houseCol);
          sheet.getRange(1, houseCol).setValue('House');
          if (sheet.getLastRow() > 1) {
            sheet.getRange(2, houseCol, sheet.getLastRow() - 1, 1).setValue('Both');
          }
          sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
        }

        if (!hasActive) {
          const refreshedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const refreshedNormalized = refreshedHeaders.map(normalizeLinkHeader_);
          const orderIndex = refreshedNormalized.indexOf('order');
          const activeCol = orderIndex === -1 ? refreshedHeaders.length + 1 : orderIndex + 2;
          sheet.insertColumnAfter(activeCol - 1);
          sheet.getRange(1, activeCol).setValue('Active');
          if (sheet.getLastRow() > 1) {
            sheet.getRange(2, activeCol, sheet.getLastRow() - 1, 1).insertCheckboxes();
          }
          sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
        }

        const finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        applyLinkColumnValidation_(sheet, finalHeaders);
      }
    }

    return sheet;
  } catch (error) {
    console.error('Error in getLinksSheet:', error);
    throw error; // Re-throw to be caught by caller
  }
}

function applyLinkColumnValidation_(sheet, headers) {
  if (!sheet || !headers || !headers.length) return;
  const headerMap = getLinkHeaderMap_(headers);
  const rowCount = Math.max(sheet.getLastRow() - 1, 1);

  if (typeof headerMap.house !== 'undefined') {
    const houseRange = sheet.getRange(2, headerMap.house + 1, rowCount, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['FOH', 'BOH', 'Both'], true)
      .setAllowInvalid(false)
      .build();
    houseRange.setDataValidation(rule);
  }

  if (typeof headerMap.active !== 'undefined') {
    const activeRange = sheet.getRange(2, headerMap.active + 1, rowCount, 1);
    activeRange.insertCheckboxes();
  }
}

/**
 * Migrate old Links sheet structure to new format
 */
function migrateLinksSheet(sheet, oldHeaders, newHeaders) {
  try {
    // Get existing data
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 1 || lastCol < 1) {
      // Empty sheet, just set headers
      sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      return;
    }

    const oldData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const now = new Date().toISOString();

    // Map old column names to indices
    const oldHeaderMap = {};
    oldHeaders.forEach((h, i) => {
      if (h) oldHeaderMap[h.toString().toLowerCase().trim()] = i;
    });

    const newData = [newHeaders];

    for (let i = 1; i < oldData.length; i++) {
      const oldRow = oldData[i];
      if (!oldRow[0] && !oldRow[1]) continue;

      const title = oldRow[oldHeaderMap['label']] || oldRow[oldHeaderMap['title']] || oldRow[0] || '';
      const url = oldRow[oldHeaderMap['url']] || oldRow[1] || '';
      const category = oldRow[oldHeaderMap['category']] || oldRow[2] || 'General';

      const row = buildLinkRow_(newHeaders, {
        link_id: generateLinkId(),
        title: title,
        url: url,
        description: '',
        icon: '📋',
        category: category,
        house: 'Both',
        display_order: i,
        status: 'Active',
        active: true
      }, {
        added_by: 'Migration',
        added_date: now
      });

      newData.push(row);
    }

    sheet.clear();
    if (newData.length > 0) {
      sheet.getRange(1, 1, newData.length, newHeaders.length).setValues(newData);
    }
    sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');
    sheet.setFrozenRows(1);

    console.log('Links sheet migrated successfully. Rows migrated: ' + (newData.length - 1));

  } catch (error) {
    console.error('Error migrating Links sheet:', error);
    // Don't throw - let the function continue with the existing sheet
  }
}

/**
 * Add default links when the sheet is first created
 */
function addDefaultLinks(sheet) {
  const defaultLinks = [
    {
      title: 'Heard Log',
      url: 'https://example.com/heard-log',
      description: 'Heard Log System',
      icon: '📋',
      category: 'Operations'
    }
  ];

  const now = new Date().toISOString();

  defaultLinks.forEach((link, index) => {
    const linkId = generateLinkId();
    const row = buildLinkRow_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0], {
      link_id: linkId,
      title: link.title,
      url: link.url,
      description: link.description,
      icon: link.icon,
      category: link.category,
      house: 'Both',
      display_order: index + 1,
      status: 'Active',
      active: true
    }, {
      added_by: 'System',
      added_date: now
    });
    sheet.appendRow(row);
  });
}

/**
 * Generate a unique link ID
 */
function generateLinkId() {
  return 'LNK' + Date.now() + Math.random().toString(36).substr(2, 5).toUpperCase();
}

/**
 * Get all links, optionally filtered by status
 * @param {boolean} includeHidden - Whether to include hidden links (default: false)
 * @returns {Object} Result with success flag and links array
 */
function getAllLinks(includeHidden) {
  // Handle undefined/null parameter (default to false)
  if (includeHidden === undefined || includeHidden === null) {
    includeHidden = false;
  }

  try {
    console.log('getAllLinks called with includeHidden:', includeHidden);
    const sheet = getLinksSheet();
    if (!sheet) {
      console.error('getLinksSheet returned null/undefined');
      return { success: false, error: 'Could not access Links sheet' };
    }

    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    console.log('Sheet info - name:', sheetName, 'lastRow:', lastRow, 'lastCol:', lastCol);

    const data = sheet.getDataRange().getValues();
    console.log('Links sheet has', data.length, 'rows');

    if (data.length <= 1) {
      console.log('No data rows found (only headers or empty)');
      return { success: true, links: [] };
    }

    const headers = data[0];
    const headerMap = getLinkHeaderMap_(headers);
    console.log('Headers:', JSON.stringify(headers));
    const links = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const link = {
        link_id: row[headerMap.link_id],
        title: row[headerMap.title],
        url: row[headerMap.url],
        description: row[headerMap.description],
        icon: row[headerMap.icon],
        category: row[headerMap.category],
        house: row[headerMap.house],
        display_order: row[headerMap.display_order],
        status: row[headerMap.status],
        active: row[headerMap.active],
        added_by: row[headerMap.added_by],
        added_date: row[headerMap.added_date],
        updated_by: row[headerMap.updated_by],
        updated_date: row[headerMap.updated_date]
      };

      const titleStr = String(link.title || '').trim();
      const urlStr = String(link.url || '').trim();
      const statusStr = String(link.status || '').trim();
      const activeFlag = typeof link.active !== 'undefined' ? parseActiveValue_(link.active) : (statusStr ? statusStr === 'Active' : true);

      if (!titleStr || !urlStr || urlStr.toLowerCase().startsWith('[url')) {
        continue;
      }

      if (!includeHidden && !activeFlag) {
        continue;
      }

      if (statusStr === 'Deleted') {
        continue;
      }

      if (link.display_order instanceof Date) {
        link.display_order = i;
      } else if (typeof link.display_order !== 'number') {
        const parsedOrder = parseInt(String(link.display_order || ''), 10);
        link.display_order = Number.isNaN(parsedOrder) ? i : parsedOrder;
      }

      link.house = String(link.house || 'Both');
      link.status = activeFlag ? 'Active' : 'Hidden';
      link.active = activeFlag;

      links.push(link);
    }

    // Sort by display_order
    links.sort((a, b) => (a.display_order || 999) - (b.display_order || 999));

    console.log('Returning', links.length, 'links');
    if (links.length > 0) {
      console.log('First link:', JSON.stringify(links[0]));
    }

    return { success: true, links: links };
  } catch (error) {
    console.error('Error getting links:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Fallback link fetch using display values to avoid serialization issues.
 * Uses getDisplayValues() to ensure all values are strings.
 *
 * @param {boolean} includeHidden
 * @returns {Object} Result with success flag and links array
 */
function getAllLinksDisplayValues(includeHidden) {
  try {
    const sheet = getLinksSheet();
    if (!sheet) {
      return { success: false, error: 'Could not access Links sheet', links: [] };
    }

    const data = sheet.getDataRange().getDisplayValues();
    if (!data || data.length <= 1) {
      return { success: true, links: [] };
    }

    const headers = data[0].map(h => (h || '').toString().trim());
    const headerMap = getLinkHeaderMap_(headers);
    const links = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const link = {
        link_id: row[headerMap.link_id],
        title: row[headerMap.title],
        url: row[headerMap.url],
        description: row[headerMap.description],
        icon: row[headerMap.icon],
        category: row[headerMap.category],
        house: row[headerMap.house],
        display_order: row[headerMap.display_order],
        status: row[headerMap.status],
        active: row[headerMap.active]
      };

      const titleStr = String(link.title || '').trim();
      const urlStr = String(link.url || '').trim();
      const statusStr = String(link.status || '').trim();
      const activeFlag = typeof link.active !== 'undefined' ? parseActiveValue_(link.active) : (statusStr ? statusStr === 'Active' : true);

      if (!titleStr || !urlStr || urlStr.toLowerCase().startsWith('[url')) {
        continue;
      }

      if (!includeHidden && !activeFlag) {
        continue;
      }
      if (statusStr === 'Deleted') {
        continue;
      }

      if (link.display_order !== undefined && link.display_order !== null) {
        const parsed = parseInt(link.display_order, 10);
        link.display_order = Number.isNaN(parsed) ? 0 : parsed;
      }

      link.house = String(link.house || 'Both');
      link.status = activeFlag ? 'Active' : 'Hidden';
      link.active = activeFlag;

      links.push(link);
    }

    links.sort((a, b) => (a.display_order || 999) - (b.display_order || 999));
    return { success: true, links: links };
  } catch (error) {
    console.error('Error in getAllLinksDisplayValues:', error);
    return { success: false, error: error.toString(), links: [] };
  }
}

/**
 * Get active links only (for Quick Links panel)
 * @returns {Object} Result with success flag and links array
 */
function getActiveLinks() {
  console.log('getActiveLinks called');
  try {
    var result = getAllLinks(false);
    console.log('getActiveLinks returning:', result ? 'valid result with ' + (result.links ? result.links.length : 0) + ' links' : 'null');
    if (!result) {
      return { success: false, error: 'getAllLinks returned null', links: [] };
    }
    return result;
  } catch (error) {
    console.error('getActiveLinks error:', error);
    return { success: false, error: error.toString(), links: [] };
  }
}

/**
 * Client-safe wrapper with session validation and EXPLICIT serialization.
 * RENAMED to avoid conflicts
 *
 * @param {boolean} includeHidden
 * @param {string} token - Session token from client
 * @returns {Object} Result with success flag and links array
 */
function api_getLinksList(includeHidden, token) {
  try {
    console.log('api_getLinksList called');
    
    // Check if Code.gs functions are available
    if (typeof getCurrentRole !== 'function') {
       console.error('CRITICAL: getCurrentRole function is missing!');
       return { success: false, error: 'Server configuration error: getCurrentRole missing' };
    }

    let session;
    try {
      session = getCurrentRole(token);
    } catch (e) {
      console.error('Error calling getCurrentRole:', e);
      return { success: false, error: 'Session validation failed: ' + e.toString() };
    }

    if (!session || !session.authenticated) {
      console.log('Session expired in api_getLinksList');
      return { success: false, sessionExpired: true, error: 'Session expired' };
    }
    
    // Call getAllLinks directly
    let result;
    try {
      result = getAllLinks(includeHidden === true);
    } catch (e) {
      console.error('Error calling getAllLinks:', e);
      return { success: false, error: 'Failed to retrieve links: ' + e.toString(), links: [] };
    }
    
    if (!result) {
      console.error('getAllLinks returned null/undefined');
      result = { success: false, error: 'Internal error: No data returned', links: [] };
    }

    // Fallback to display values if primary fetch failed
    if (!result.success) {
      console.warn('Primary link fetch failed, trying display-value fallback:', result.error || '');
      const fallback = getAllLinksDisplayValues(includeHidden === true);
      if (fallback && fallback.success) {
        result = fallback;
      } else {
        return { success: false, error: fallback.error || result.error || 'Failed to load links', links: [] };
      }
    }

    // SANITIZE: Explicitly convert complex objects to primitives
    // This prevents "null" response due to serialization failures
    if (result.success && result.links) {
      const cleanLinks = result.links.map(link => {
        const clean = {};
        try {
            // Copy only known properties with explicit conversion
            clean.link_id = String(link.link_id || '');
            clean.title = String(link.title || '');
            clean.url = String(link.url || '');
            clean.description = String(link.description || '');
            clean.icon = String(link.icon || '');
            clean.category = String(link.category || '');
            clean.house = String(link.house || 'Both');
            clean.display_order = Number(link.display_order) || 0;
            clean.status = String(link.status || '');
            clean.active = typeof link.active === 'boolean' ? link.active : parseActiveValue_(link.active);
            clean.added_by = String(link.added_by || '');
            
            // Handle dates specifically
            if (link.added_date instanceof Date) {
              clean.added_date = link.added_date.toISOString();
            } else {
              clean.added_date = String(link.added_date || '');
            }
        } catch (err) {
            console.error('Error sanitizing link:', err);
            // Return a minimal safe object if sanitization fails for a specific link
            clean.link_id = String(link.link_id || 'ERROR');
            clean.title = 'Error loading link';
        }
        return clean;
      });
      
      console.log('Sanitized ' + cleanLinks.length + ' links');
      return { success: true, links: cleanLinks };
    }
    
    // If result was not success or had no links, ensure we return a safe object
    return { 
        success: Boolean(result.success), 
        error: String(result.error || ''),
        links: []
    };

  } catch (error) {
    console.error('api_getLinksList global error:', error);
    // Log to sheet if possible
    try {
      logLinkAction('ERROR', 'getLinksList', error.toString(), 'System');
    } catch (e) {
      console.error('Failed to log error to sheet:', e);
    }
    return { success: false, error: 'Global error: ' + error.toString(), links: [] };
  }
}

/**
 * Validate URL format
 * @param {string} url - The URL to validate
 * @returns {boolean} True if valid URL
 */
function validateUrl(url) {
  if (!url) return false;

  url = url.trim();

  // Must start with http:// or https://
  if (!url.startsWith('http://') && !url.startsWith('https://')) {
    return false;
  }

  // Comprehensive URL pattern that works in Google Apps Script
  // Allows: domain names, paths, query strings, fragments, ports, etc.
  const urlPattern = /^https?:\/\/[a-zA-Z0-9][-a-zA-Z0-9]*(\.[a-zA-Z0-9][-a-zA-Z0-9]*)+(:\d+)?(\/[^\s]*)?$/;

  return urlPattern.test(url);
}

/**
 * Normalize URL (add https:// if missing)
 * @param {string} url - The URL to normalize
 * @returns {string} Normalized URL
 */
function normalizeUrl(url) {
  if (!url) return '';

  url = url.trim();

  if (!url.startsWith('http://') && !url.startsWith('https://')) {
    url = 'https://' + url;
  }

  return url;
}

/**
 * Add a new link
 * @param {Object} linkData - The link data
 * @param {string} addedByDirector - Name of the director adding the link
 * @returns {Object} Result with success flag and new link ID
 */
function addLink(linkData, addedByDirector) {
  try {
    console.log('addLink called with:', JSON.stringify(linkData), 'by:', addedByDirector);

    // Validate required fields
    if (!linkData || typeof linkData !== 'object') {
      return { success: false, error: 'Invalid link data provided' };
    }

    if (!linkData.title || !linkData.title.trim()) {
      return { success: false, error: 'Link title is required' };
    }

    if (!linkData.url || !linkData.url.trim()) {
      return { success: false, error: 'Link URL is required' };
    }

    // Normalize and validate URL
    const normalizedUrl = normalizeUrl(linkData.url);
    if (!validateUrl(normalizedUrl)) {
      return { success: false, error: 'Invalid URL format' };
    }

    // Check for duplicate URL within the same house
    const existingLinks = getAllLinks(true);
    if (existingLinks.success) {
      const incomingHouse = String(linkData.house || 'Both').toUpperCase();
      const duplicate = existingLinks.links.find(
        link => link.url && normalizedUrl &&
                link.url.toLowerCase() === normalizedUrl.toLowerCase() &&
                String(link.house || 'Both').toUpperCase() === incomingHouse &&
                link.status !== 'Deleted'
      );
      if (duplicate) {
        return { success: false, error: 'A link with this URL already exists: ' + duplicate.title };
      }
    }

    const sheet = getLinksSheet();
    const linkId = generateLinkId();
    const now = new Date().toISOString();

    const allLinks = getAllLinks(true);
    const maxOrder = (allLinks.success && allLinks.links)
      ? allLinks.links.reduce((max, link) => Math.max(max, link.display_order || 0), 0)
      : 0;

    const houseValue = String(linkData.house || 'Both');
    const normalizedHouse = ['FOH', 'BOH', 'Both'].includes(houseValue) ? houseValue : 'Both';
    const requestedOrder = parseInt(String(linkData.display_order || ''), 10);
    const displayOrder = Number.isNaN(requestedOrder) ? maxOrder + 1 : Math.max(requestedOrder, 1);
    const activeFlag = typeof linkData.active === 'boolean'
      ? linkData.active
      : String(linkData.status || 'Active').toLowerCase() === 'active';

    const row = buildLinkRow_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0], {
      link_id: linkId,
      title: linkData.title.trim(),
      url: normalizedUrl,
      description: (linkData.description || '').trim(),
      icon: linkData.icon || 'link',
      category: linkData.category || 'General',
      house: normalizedHouse,
      display_order: displayOrder,
      status: activeFlag ? 'Active' : 'Hidden',
      active: activeFlag
    }, {
      added_by: addedByDirector || 'Unknown',
      added_date: now
    });

    sheet.appendRow(row);

    // Log the action
    logLinkAction('ADD', linkId, linkData.title, addedByDirector);

    return {
      success: true,
      link_id: linkId,
      message: 'Link added successfully'
    };
  } catch (error) {
    console.error('Error adding link:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Client-safe wrapper for addLink with session validation.
 * RENAMED to avoid conflicts
 */
function api_addLink(linkData, token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true, error: 'Session expired' };
    }
    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Access denied' };
    }
    return addLink(linkData, session.role);
  } catch (error) {
    console.error('Error in api_addLink:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Update an existing link
 * @param {string} linkId - The link ID to update
 * @param {Object} updateData - The fields to update
 * @param {string} updatedByDirector - Name of the director making the update
 * @returns {Object} Result with success flag
 */
function updateLink(linkId, updateData, updatedByDirector) {
  try {
    if (!linkId) {
      return { success: false, error: 'Link ID is required' };
    }

    const sheet = getLinksSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const headerMap = getLinkHeaderMap_(headers);

    // Find the link row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === linkId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'Link not found' };
    }

    // Validate URL if being updated
    if (updateData.url) {
      const normalizedUrl = normalizeUrl(updateData.url);
      if (!validateUrl(normalizedUrl)) {
        return { success: false, error: 'Invalid URL format' };
      }

      // Check for duplicate URL (excluding current link) within the same house
      const existingLinks = getAllLinks(true);
      if (existingLinks.success) {
        const existingHouse = headerMap.house != null ? data[rowIndex][headerMap.house] : 'Both';
        const incomingHouse = String(
          typeof updateData.house !== 'undefined' ? updateData.house : existingHouse || 'Both'
        ).toUpperCase();
        const duplicate = existingLinks.links.find(
          link => link.link_id !== linkId &&
                  link.url && normalizedUrl &&
                  link.url.toLowerCase() === normalizedUrl.toLowerCase() &&
                  String(link.house || 'Both').toUpperCase() === incomingHouse &&
                  link.status !== 'Deleted'
        );
        if (duplicate) {
          return { success: false, error: 'A link with this URL already exists: ' + duplicate.title };
        }
      }

      updateData.url = normalizedUrl;
    }

    // Update the fields
    const now = new Date().toISOString();
    const currentRow = data[rowIndex];

    // Update allowed fields
    const allowedFields = ['title', 'url', 'description', 'icon', 'category', 'house', 'display_order', 'status', 'active'];

    if (updateData.hasOwnProperty('house')) {
      const houseValue = String(updateData.house || 'Both');
      updateData.house = ['FOH', 'BOH', 'Both'].includes(houseValue) ? houseValue : 'Both';
    }

    allowedFields.forEach(field => {
      if (!updateData.hasOwnProperty(field)) return;
      const colIndex = headerMap[field];
      if (typeof colIndex === 'undefined') return;
      let value = updateData[field];

      if (field === 'display_order') {
        const parsedOrder = parseInt(String(value || ''), 10);
        value = Number.isNaN(parsedOrder) ? currentRow[colIndex] : parsedOrder;
      }

      if (field === 'active') {
        const activeFlag = parseActiveValue_(value);
        currentRow[colIndex] = activeFlag ? true : false;
        if (typeof headerMap.status !== 'undefined') {
          currentRow[headerMap.status] = activeFlag ? 'Active' : 'Hidden';
        }
        return;
      }

      if (field === 'status') {
        const activeFlag = String(value || '').toLowerCase() === 'active';
        if (typeof headerMap.active !== 'undefined') {
          currentRow[headerMap.active] = activeFlag ? true : false;
        }
      }

      if (typeof value === 'string') {
        value = value.trim();
      }

      currentRow[colIndex] = value;
    });

    // Update metadata
    if (typeof headerMap.updated_by !== 'undefined') {
      currentRow[headerMap.updated_by] = updatedByDirector || 'Unknown';
    }
    if (typeof headerMap.updated_date !== 'undefined') {
      currentRow[headerMap.updated_date] = now;
    }

    // Write back to sheet
    sheet.getRange(rowIndex + 1, 1, 1, currentRow.length).setValues([currentRow]);

    // Log the action
    logLinkAction('UPDATE', linkId, updateData.title || data[rowIndex][1], updatedByDirector);

    return {
      success: true,
      message: 'Link updated successfully'
    };
  } catch (error) {
    console.error('Error updating link:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Client-safe wrapper for updateLink with session validation.
 * RENAMED to avoid conflicts
 */
function api_updateLink(linkId, updateData, token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true, error: 'Session expired' };
    }
    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Access denied' };
    }
    return updateLink(linkId, updateData, session.role);
  } catch (error) {
    console.error('Error in api_updateLink:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Delete a link (soft delete - marks as Deleted)
 * @param {string} linkId - The link ID to delete
 * @param {string} deletedByDirector - Name of the director deleting the link
 * @returns {Object} Result with success flag
 */
function deleteLink(linkId, deletedByDirector) {
  try {
    if (!linkId) {
      return { success: false, error: 'Link ID is required' };
    }

    const sheet = getLinksSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Find the link row
    let rowIndex = -1;
    let linkTitle = '';
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === linkId) {
        rowIndex = i;
        linkTitle = data[i][1];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'Link not found' };
    }

    // Soft delete by setting status to 'Deleted'
    const now = new Date().toISOString();
    const statusCol = headerMap.status;
    const activeCol = headerMap.active;
    const updatedByCol = headerMap.updated_by;
    const updatedDateCol = headerMap.updated_date;

    if (typeof statusCol !== 'undefined') {
      sheet.getRange(rowIndex + 1, statusCol + 1).setValue('Deleted');
    } else if (typeof activeCol !== 'undefined') {
      sheet.getRange(rowIndex + 1, activeCol + 1).setValue(false);
    } else {
      sheet.deleteRow(rowIndex + 1);
      return { success: true, message: 'Link deleted successfully' };
    }

    if (typeof updatedByCol !== 'undefined') {
      sheet.getRange(rowIndex + 1, updatedByCol + 1).setValue(deletedByDirector || 'Unknown');
    }
    if (typeof updatedDateCol !== 'undefined') {
      sheet.getRange(rowIndex + 1, updatedDateCol + 1).setValue(now);
    }

    // Log the action
    logLinkAction('DELETE', linkId, linkTitle, deletedByDirector);

    return {
      success: true,
      message: 'Link deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting link:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Client-safe wrapper for deleteLink with session validation.
 * RENAMED to avoid conflicts
 */
function api_deleteLink(linkId, token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true, error: 'Session expired' };
    }
    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Access denied' };
    }
    return deleteLink(linkId, session.role);
  } catch (error) {
    console.error('Error in api_deleteLink:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Toggle link status between Active and Hidden
 * @param {string} linkId - The link ID
 * @param {string} newStatus - New status ('Active' or 'Hidden')
 * @param {string} updatedByDirector - Name of the director making the change
 * @returns {Object} Result with success flag
 */
function toggleLinkStatus(linkId, newStatus, updatedByDirector) {
  try {
    if (!linkId) {
      return { success: false, error: 'Link ID is required' };
    }

    if (!['Active', 'Hidden'].includes(newStatus)) {
      return { success: false, error: 'Invalid status. Must be Active or Hidden' };
    }

    return updateLink(linkId, { status: newStatus }, updatedByDirector);
  } catch (error) {
    console.error('Error toggling link status:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Client-safe wrapper for toggleLinkStatus with session validation.
 * RENAMED to avoid conflicts
 */
function api_toggleLinkStatus(linkId, newStatus, token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true, error: 'Session expired' };
    }
    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Access denied' };
    }
    return toggleLinkStatus(linkId, newStatus, session.role);
  } catch (error) {
    console.error('Error in api_toggleLinkStatus:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Reorder links by updating their display_order
 * @param {Array} linkIdsInOrder - Array of link IDs in the new order
 * @param {string} reorderedByDirector - Name of the director making the change
 * @returns {Object} Result with success flag
 */
function reorderLinks(linkIdsInOrder, reorderedByDirector) {
  try {
    if (!linkIdsInOrder || !Array.isArray(linkIdsInOrder) || linkIdsInOrder.length === 0) {
      return { success: false, error: 'Link IDs array is required' };
    }

    const sheet = getLinksSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const headerMap = getLinkHeaderMap_(headers);

    const displayOrderCol = headerMap.display_order;
    const updatedByCol = headerMap.updated_by;
    const updatedDateCol = headerMap.updated_date;

    if (typeof displayOrderCol === 'undefined') {
      return { success: false, error: 'display_order column not found' };
    }

    const now = new Date().toISOString();

    // Update each link's display_order
    linkIdsInOrder.forEach((linkId, index) => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === linkId) {
          sheet.getRange(i + 1, displayOrderCol + 1).setValue(index + 1);
          if (typeof updatedByCol !== 'undefined') {
            sheet.getRange(i + 1, updatedByCol + 1).setValue(reorderedByDirector || 'Unknown');
          }
          if (typeof updatedDateCol !== 'undefined') {
            sheet.getRange(i + 1, updatedDateCol + 1).setValue(now);
          }
          break;
        }
      }
    });

    // Log the action
    logLinkAction('REORDER', 'multiple', 'Links reordered', reorderedByDirector);

    return {
      success: true,
      message: 'Links reordered successfully'
    };
  } catch (error) {
    console.error('Error reordering links:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Client-safe wrapper for reorderLinks with session validation.
 * RENAMED to avoid conflicts
 */
function api_reorderLinks(linkIdsInOrder, token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true, error: 'Session expired' };
    }
    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Access denied' };
    }
    return reorderLinks(linkIdsInOrder, session.role);
  } catch (error) {
    console.error('Error in api_reorderLinks:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Get available icon options
 * @returns {Array} Array of icon objects with name and label
 */
function getAvailableIcons() {
  return [
    { name: 'link', label: 'Link' },
    { name: 'external-link-alt', label: 'External Link' },
    { name: 'globe', label: 'Globe' },
    { name: 'graduation-cap', label: 'Graduation Cap' },
    { name: 'book', label: 'Book' },
    { name: 'file-alt', label: 'Document' },
    { name: 'calendar', label: 'Calendar' },
    { name: 'clock', label: 'Clock' },
    { name: 'users', label: 'Users' },
    { name: 'user', label: 'User' },
    { name: 'comments', label: 'Comments' },
    { name: 'envelope', label: 'Email' },
    { name: 'phone', label: 'Phone' },
    { name: 'video', label: 'Video' },
    { name: 'chart-bar', label: 'Chart' },
    { name: 'cog', label: 'Settings' },
    { name: 'tools', label: 'Tools' },
    { name: 'clipboard-list', label: 'Clipboard' },
    { name: 'tasks', label: 'Tasks' },
    { name: 'utensils', label: 'Food' },
    { name: 'store', label: 'Store' },
    { name: 'dollar-sign', label: 'Dollar' },
    { name: 'credit-card', label: 'Credit Card' },
    { name: 'briefcase', label: 'Briefcase' },
    { name: 'building', label: 'Building' },
    { name: 'map-marker-alt', label: 'Location' },
    { name: 'info-circle', label: 'Info' },
    { name: 'question-circle', label: 'Help' },
    { name: 'exclamation-triangle', label: 'Warning' },
    { name: 'check-circle', label: 'Check' },
    { name: 'star', label: 'Star' },
    { name: 'heart', label: 'Heart' }
  ];
}

/**
 * Get available category options
 * @returns {Array} Array of category names
 */
function getAvailableCategories() {
  return [
    'Training',
    'Communication',
    'Operations',
    'HR',
    'Finance',
    'Scheduling',
    'Safety',
    'Marketing',
    'General',
    'Other'
  ];
}

/**
 * Log link management actions to the Logs sheet
 * @param {string} action - The action type (ADD, UPDATE, DELETE, REORDER)
 * @param {string} linkId - The link ID
 * @param {string} details - Additional details
 * @param {string} performedBy - Who performed the action
 */
function logLinkAction(action, linkId, details, performedBy) {
  try {
    console.log('logLinkAction called:', action, linkId);
    const ss = SpreadsheetApp.openById(LINK_SPREADSHEET_ID);
    let logsSheet = ss.getSheetByName('Logs');

    if (!logsSheet) {
      logsSheet = ss.insertSheet('Logs');
      logsSheet.getRange(1, 1, 1, 5).setValues([['timestamp', 'action_type', 'entity_id', 'details', 'performed_by']]);
      logsSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      logsSheet.setFrozenRows(1);
    }

    const now = new Date().toISOString();
    logsSheet.appendRow([now, 'LINK_' + action, linkId, details, performedBy || 'Unknown']);

  } catch (error) {
    console.error('Error logging link action:', error);
    // Don't throw - logging failures shouldn't break the main operation
  }
}

/**
 * Get link by ID
 * @param {string} linkId - The link ID
 * @returns {Object} Result with success flag and link data
 */
function getLinkById(linkId) {
  try {
    if (!linkId) {
      return { success: false, error: 'Link ID is required' };
    }

    const allLinks = getAllLinks(true);
    if (!allLinks.success) {
      return allLinks;
    }

    const link = allLinks.links.find(l => l.link_id === linkId);

    if (!link) {
      return { success: false, error: 'Link not found' };
    }

    return { success: true, link: link };
  } catch (error) {
    console.error('Error getting link by ID:', error);
    return { success: false, error: error.toString() };
  }
}


// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Diagnostic function to check Links sheet structure
 * Run this from Apps Script editor to verify the sheet is set up correctly
 */
function diagnoseLinkSheet() {
  console.log('=== Links Sheet Diagnostic ===\n');

  try {
    console.log('1. Checking SPREADSHEET_ID...');
    console.log('   LINK_SPREADSHEET_ID:', LINK_SPREADSHEET_ID);

    console.log('\n2. Opening spreadsheet...');
    const ss = SpreadsheetApp.openById(LINK_SPREADSHEET_ID);
    console.log('   Spreadsheet name:', ss.getName());

    console.log('\n3. Getting Links sheet...');
    const sheet = ss.getSheetByName('Links');
    if (!sheet) {
      console.log('   ERROR: Links sheet does not exist!');
      console.log('   Creating Links sheet via getLinksSheet()...');
      const newSheet = getLinksSheet();
      console.log('   Created sheet:', newSheet.getName());
    } else {
      console.log('   Links sheet exists');
    }

    console.log('\n4. Reading sheet data...');
    const linksSheet = ss.getSheetByName('Links');
    const lastRow = linksSheet.getLastRow();
    const lastCol = linksSheet.getLastColumn();
    console.log('   Last row:', lastRow);
    console.log('   Last column:', lastCol);

    if (lastRow > 0 && lastCol > 0) {
      const allData = linksSheet.getRange(1, 1, lastRow, lastCol).getValues();
      console.log('\n5. Headers (row 1):');
      console.log('   ', JSON.stringify(allData[0]));

      console.log('\n6. Data rows:');
      for (let i = 1; i < allData.length; i++) {
        console.log('   Row', i + 1, ':', JSON.stringify(allData[i]));
      }
    } else {
      console.log('   Sheet is empty');
    }

    console.log('\n7. Testing getAllLinks()...');
    const result = getAllLinks(true);
    console.log('   Success:', result.success);
    console.log('   Links count:', result.links ? result.links.length : 0);
    if (result.links && result.links.length > 0) {
      console.log('   First link:', JSON.stringify(result.links[0]));
    }
    if (result.error) {
      console.log('   Error:', result.error);
    }

    console.log('\n=== Diagnostic Complete ===');
    return { success: true, message: 'Diagnostic complete - check logs' };

  } catch (error) {
    console.error('Diagnostic error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Test function to verify link management functionality
 */
function testLinkManagement() {
  console.log('=== Testing Link Management ===');

  // Test 1: Get all links
  console.log('\n1. Getting all links...');
  const allLinks = getAllLinks(true);
  console.log('Links found:', allLinks.links ? allLinks.links.length : 0);

  // Test 2: Add a new link
  console.log('\n2. Adding a new link...');
  const newLink = addLink({
    title: 'Test Link',
    url: 'https://test.example.com',
    description: 'This is a test link',
    icon: 'star',
    category: 'General'
  }, 'Test Director');
  console.log('Add result:', newLink);

  if (newLink.success) {
    // Test 3: Get the new link
    console.log('\n3. Getting the new link...');
    const fetchedLink = getLinkById(newLink.link_id);
    console.log('Fetched link:', fetchedLink);

    // Test 4: Update the link
    console.log('\n4. Updating the link...');
    const updateResult = updateLink(newLink.link_id, {
      title: 'Updated Test Link',
      description: 'This description was updated'
    }, 'Test Director');
    console.log('Update result:', updateResult);

    // Test 5: Toggle status
    console.log('\n5. Toggling link status to Hidden...');
    const toggleResult = toggleLinkStatus(newLink.link_id, 'Hidden', 'Test Director');
    console.log('Toggle result:', toggleResult);

    // Test 6: Delete the link
    console.log('\n6. Deleting the link...');
    const deleteResult = deleteLink(newLink.link_id, 'Test Director');
    console.log('Delete result:', deleteResult);
  }

  // Test 7: Get available icons
  console.log('\n7. Getting available icons...');
  const icons = getAvailableIcons();
  console.log('Icons available:', icons.length);

  // Test 8: Get available categories
  console.log('\n8. Getting available categories...');
  const categories = getAvailableCategories();
  console.log('Categories:', categories);

  console.log('\n=== Link Management Tests Complete ===');
}

/**
 * Test URL validation
 */
function testUrlValidation() {
  console.log('=== Testing URL Validation ===');

  const testUrls = [
    'https://example.com',
    'http://example.com',
    'example.com',
    'www.example.com',
    'https://subdomain.example.com/path/to/page',
    'invalid',
    '',
    'ftp://example.com',
    'https://example.com/path?query=1&other=2'
  ];

  testUrls.forEach(url => {
    const normalized = normalizeUrl(url);
    const isValid = validateUrl(normalized);
    console.log(`URL: "${url}" -> Normalized: "${normalized}" -> Valid: ${isValid}`);
  });

  console.log('=== URL Validation Tests Complete ===');
}
