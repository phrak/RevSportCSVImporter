/********************************************
* SCOUTHUB CSV IMPORT TRANSFORMS MODULE v4.0
*
* SUMMARY:
* Data transformation utilities for ScoutHub Member Data Import
*
* Functional Features:
* - Configurable target sheet (from Import Config!B2)
* - Easily update column names and config in the CONFIG section
*
* Single-click batch processing of all transforms
* - Parent name splitting (first/last) for both parents
* - Automatic population of Preferred Name field
* - Sorting by Membership Number column
* - Batch normalization of phone numbers to Australian formats
* - Invalid phone numbers are highlighted (red cell, white text)
* - Mobile Number & Email Address de-duplication
*
* TESTING & ERROR HANDLING
* - Error handling with user-friendly alerts and logging
* - Built-in test utility for phone number normalization
* - Performance monitoring for long-running operations
********************************************/

const TransformsModule = (function() {
  'use strict';

  /*********************************
   * CORE CONFIGURATION (PERSISTENT)
   *********************************/

  /**
   * Loads configuration from PropertyService, or uses defaults if not set.
   * Implements caching for better performance.
   * @returns {Object} The configuration object.
   */
  function _loadConfig() {
    // First check cache for config to reduce property service calls
    const cache = CacheService.getScriptCache();
    const cachedConfig = cache.get('TRANSFORM_CONFIG');
    
    if (cachedConfig) {
      try {
        return JSON.parse(cachedConfig);
      } catch (e) {
        Logger.log('Failed to parse cached config, trying stored config.');
      }
    }
    
    // If not in cache, try script properties
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty('TRANSFORM_CONFIG');
    
    if (stored) {
      try {
        const config = JSON.parse(stored);
        // Cache for 10 minutes to improve performance
        cache.put('TRANSFORM_CONFIG', stored, 600);
        return config;
      } catch (e) {
        Logger.log('Failed to parse stored config, using defaults.');
      }
    }

    // Default config
    const defaultConfig = {
      CONFIG_SHEET: 'Import Config',
      TARGET_SHEET_CELL: 'B2',
      PHONE_COLUMNS: ['Member Mobile', 'Parent1_Mobile', 'Parent2_Mobile', 'Home_Number'],
      PARENT1_FNAME: 'Parent1_FName',
      PARENT1_LNAME: 'Parent1_LName',
      PARENT2_FNAME: 'Parent2_FName',
      PARENT2_LNAME: 'Parent2_LName',
      PREFERRED_NAME: 'Preferred Name',
      FIRST_NAME: 'First Name',
      MEMBER_NUMBER: 'Membership Number',
      MEMBER_MOBILE: 'Member Mobile',
      PARENT1_MOBILE: 'Parent1_Mobile',
      PARENT2_MOBILE: 'Parent2_Mobile',
      MEMBER_EMAIL: 'Member_Email',
      ADDITIONAL_EMAILS: 'Additional Email Addresses',
      PARENT1_EMAIL: 'Parent1_Email',
      PARENT2_EMAIL: 'Parent2_Email'
    };
    
    // Cache default config for future use
    cache.put('TRANSFORM_CONFIG', JSON.stringify(defaultConfig), 600);
    return defaultConfig;
  }

  /**
   * Saves configuration to PropertyService and updates cache.
   * @param {Object} config The configuration object to save.
   */
  function _saveConfig(config) {
    const configString = JSON.stringify(config);
    PropertiesService.getScriptProperties().setProperty('TRANSFORM_CONFIG', configString);
    // Update cache for faster access
    CacheService.getScriptCache().put('TRANSFORM_CONFIG', configString, 600);
  }

  let CONFIG = _loadConfig();

  /**
   * Allows persistent updating of configuration keys.
   * @param {string} key The config property to update.
   * @param {*} value The value to set.
   */
  function updateConfig(key, value) {
    CONFIG[key] = value;
    _saveConfig(CONFIG);
    // Clear related caches to ensure fresh data
    _clearRelatedCaches();
  }

  /**
   * Clears caches related to configuration changes.
   * Ensures fresh data when config is updated.
   */
  function _clearRelatedCaches() {
    const cache = CacheService.getScriptCache();
    const keysToDelete = ['targetSheetName', 'headerIndices', 'sheetData'];
    cache.removeAll(keysToDelete);
  }

  /**
   * Resets all persistent configuration to defaults.
   */
  function resetAllConfigs() {
    try {
      PropertiesService.getScriptProperties().deleteAllProperties();
      PropertiesService.getDocumentProperties().deleteAllProperties();
      PropertiesService.getUserProperties().deleteAllProperties();
      
      // Clear all caches
      CacheService.getScriptCache().removeAll([
        'TRANSFORM_CONFIG', 'targetSheetName', 'headerIndices', 'sheetData'
      ]);
      
      CONFIG = _loadConfig(); // Reload default config
      SpreadsheetApp.getUi().alert('All configurations reset to defaults');
    } catch (e) {
      _handleError('Config reset failed', e);
    }
  }

  /*********************************
   * PERFORMANCE MONITORING
   *********************************/
  
  /**
   * Simple performance monitoring utility.
   * Tracks execution time of operations.
   */
  const Performance = {
    timers: {},
    
    /**
     * Starts a timer for the specified operation.
     * @param {string} operation Name of the operation to time
     */
    start: function(operation) {
      this.timers[operation] = new Date().getTime();
    },
    
    /**
     * Ends a timer and logs the duration.
     * @param {string} operation Name of the operation that was timed
     * @returns {number} Duration in milliseconds
     */
    end: function(operation) {
      if (!this.timers[operation]) return 0;
      
      const duration = new Date().getTime() - this.timers[operation];
      Logger.log(`Operation "${operation}" completed in ${duration}ms`);
      delete this.timers[operation];
      return duration;
    }
  };

  /*********************************
   * SHEET/CONFIG HELPERS
   *********************************/

  /**
   * Gets the target sheet name from configuration.
   * Uses caching to minimize expensive spreadsheet operations.
   * @returns {string} The target sheet name.
   */
  function _getMappedMembersSheetName() {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('targetSheetName');
    
    if (cached) return cached;
    
    try {
      const configSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.CONFIG_SHEET);
      if (!configSheet) throw new TransformError('Config sheet not found', 'MISSING_CONFIG_SHEET');
      
      const sheetName = configSheet.getRange(CONFIG.TARGET_SHEET_CELL).getValue();
      if (!sheetName) throw new TransformError('Target sheet name not configured in B2', 'MISSING_TARGET_NAME');
      
      cache.put('targetSheetName', sheetName, 600); // Cache for 10 minutes
      return sheetName;
    } catch (e) {
      if (e instanceof TransformError) throw e;
      throw new TransformError('Configuration retrieval failed', 'CONFIG_ERROR');
    }
  }

  /**
   * Gets the target sheet object.
   * @returns {Sheet} The Google Sheets object.
   */
  function _getSheet() {
    const sheetName = _getMappedMembersSheetName();
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    
    if (!sheet) throw new TransformError(`Target sheet '${sheetName}' not found`, 'MISSING_TARGET_SHEET');
    return sheet;
  }

  /**
   * Gets all data from the target sheet including headers.
   * Uses caching for better performance on repeated access.
   * @param {boolean} [forceRefresh=false] Force refresh of cached data
   * @returns {Object} Object containing headers, data, and sheet reference.
   */
  function _getSheetData(forceRefresh = false) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'sheetData';
    
    if (!forceRefresh) {
      const cachedData = cache.get(cacheKey);
      if (cachedData) {
        try {
          const parsed = JSON.parse(cachedData);
          // Return with sheet reference which can't be cached
          return { 
            headers: parsed.headers, 
            data: parsed.data, 
            sheet: _getSheet() 
          };
        } catch (e) {
          Logger.log('Failed to parse cached sheet data');
        }
      }
    }
    
    // If no cache or forced refresh, get from sheet
    Performance.start('getSheetData');
    const sheet = _getSheet();
    const data = sheet.getDataRange().getValues();
    const result = { 
      headers: data[0], 
      data: data.slice(1), 
      sheet: sheet 
    };
    
    // Cache the data (cannot cache sheet reference)
    const cacheData = JSON.stringify({
      headers: result.headers,
      data: result.data
    });
    
    cache.put(cacheKey, cacheData, 600); // Cache for 10 minutes
    Performance.end('getSheetData');
    
    return result;
  }

  /**
   * Updates sheet data in batch to minimize API calls.
   * @param {Array<Array>} data The data to write to the sheet.
   */
  function _updateSheetData(data) {
    if (!data || data.length === 0) return;
    
    Performance.start('updateSheetData');
    _getSheet().getRange(2, 1, data.length, data[0].length).setValues(data);
    Performance.end('updateSheetData');
    
    // Clear cached data since it's now stale
    CacheService.getScriptCache().remove('sheetData');
  }

  /**
   * Gets column indices for specified header names.
   * Uses caching to prevent repeated header scans.
   * @param {Array<string>} headerNames Array of header names to find.
   * @returns {Object} Object with headers array and cols array of indices.
   */
  function _getHeaderIndices(headerNames) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'headerIndices';
    const cachedIndices = cache.get(cacheKey);
    
    if (cachedIndices) {
      try {
        const parsed = JSON.parse(cachedIndices);
        // Filter requested headers from cached data
        const cols = headerNames.map(h => parsed.indices[h]).filter(i => i !== undefined);
        return { headers: parsed.headers, cols: cols };
      } catch (e) {
        Logger.log('Failed to parse cached header indices');
      }
    }
    
    // Get headers if not cached
    Performance.start('getHeaderIndices');
    const headers = _getSheet().getRange(1, 1, 1, _getSheet().getLastColumn()).getValues()[0];
    
    // Create map of all header names to indices
    const indicesMap = {};
    headers.forEach((header, index) => {
      indicesMap[header] = index;
    });
    
    // Cache all header indices
    cache.put(cacheKey, JSON.stringify({ 
      headers: headers, 
      indices: indicesMap 
    }), 600);
    
    // Filter for requested headers
    const cols = headerNames.map(h => indicesMap[h]).filter(i => i !== undefined);
    Performance.end('getHeaderIndices');
    
    return { headers, cols };
  }

  /**
   * Converts a column number to letter notation (A, B, C, etc).
   * @param {number} column The column number (1-based).
   * @returns {string} The column letter(s).
   */
  function _columnToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

  /*********************************
   * PHONE PROCESSING (BATCHED)
   *********************************/

  /**
   * Normalize phone number to Australian format.
   * Handles +61, 61, missing leading zero, 8/9-digit numbers, and formats output.
   * @param {string|number} value The phone number.
   * @returns {string} The normalized phone number.
   */
  function _normalizePhoneNumber(value) {
    if (!value) return '';
    
    var cleanedValue = String(value).replace(/\s/g, '').replace(/[^+\d]/g, ''); // Remove spaces and non-digits except +
    
    // If empty after cleaning, return empty string
    if (!cleanedValue) return '';
    
    // Handle +61 and 61 formats
    if (cleanedValue.startsWith('+61')) {
      cleanedValue = '0' + cleanedValue.slice(3);
    } else if (cleanedValue.startsWith('61') && cleanedValue.length >= 10) {
      cleanedValue = '0' + cleanedValue.slice(2);
    }

    // Add leading digits for mobiles and landlines
    if (cleanedValue.length === 9 && cleanedValue.startsWith('4')) {
      cleanedValue = '0' + cleanedValue;
    } else if (cleanedValue.startsWith('2') && cleanedValue.length === 9) {
      cleanedValue = '0' + cleanedValue;
    } else if ((cleanedValue.startsWith('8') || cleanedValue.startsWith('9')) && cleanedValue.length === 8) {
      cleanedValue = '02' + cleanedValue;
    }

    // Format the number
    if (cleanedValue.startsWith('02') && cleanedValue.length === 10) {
      cleanedValue = cleanedValue.replace(/(\d{2})(\d{4})(\d{4})/, '$1 $2 $3');
    } else if (cleanedValue.length === 10 && cleanedValue.startsWith('04')) {
      cleanedValue = cleanedValue.replace(/(\d{4})(\d{3})(\d{3})/, '$1 $2 $3');
    }

    return cleanedValue;
  }

  /**
   * Validates if a phone number follows Australian format patterns.
   * @param {string} formatted The formatted phone number to validate.
   * @returns {boolean} True if the phone number is valid.
   */
  function _isValidAustralianPhone(formatted) {
    if (!formatted) return false;
    return /^04\d{2} \d{3} \d{3}$/.test(formatted) || /^0[2378] \d{4} \d{4}$/.test(formatted);
  }

  /**
   * Processes all phone columns in the sheet.
   * Optimized to minimize API calls by batching operations.
   */
  function _processPhoneColumns() {
    Performance.start('processPhoneColumns');
    const sheet = _getSheet();
    const { headers } = _getSheetData();
    
    // Get all phone column indices
    const phoneColIndices = CONFIG.PHONE_COLUMNS.map(col => headers.indexOf(col)).filter(idx => idx !== -1);
    
    if (phoneColIndices.length === 0) {
      throw new TransformError('No valid phone columns found', 'MISSING_COLUMNS');
    }
    
    // Read all data in one operation
    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    let modified = false;
    
    // Process phone numbers in memory
    for (let i = 1; i < allData.length; i++) {
      for (const colIdx of phoneColIndices) {
        const originalValue = allData[i][colIdx];
        if (!originalValue) continue;
        
        const normalizedValue = _normalizePhoneNumber(originalValue);
        if (normalizedValue !== originalValue) {
          allData[i][colIdx] = normalizedValue;
          modified = true;
        }
      }
    }

    // Only update sheet if changes were made
    if (modified) {
      dataRange.setValues(allData);
      // Clear cache since data changed
      CacheService.getScriptCache().remove('sheetData');
    }

    // Apply conditional formatting to highlight invalid numbers
    _applyPhoneConditionalFormatting(sheet, phoneColIndices);
    
    Performance.end('processPhoneColumns');
  }

  /**
   * Applies conditional formatting to highlight invalid phone numbers.
   * @param {Sheet} sheet The sheet to apply formatting to.
   * @param {Array<number>} phoneColIndices Indices of columns containing phone numbers.
   */
  function _applyPhoneConditionalFormatting(sheet, phoneColIndices) {
    Performance.start('applyPhoneConditionalFormatting');
    
    // Batch process the conditional formatting rules
    let rules = sheet.getConditionalFormatRules();
    const newRules = [];
    
    // Remove existing rules for phone columns
    phoneColIndices.forEach(colIdx => {
      const columnLetter = _columnToLetter(colIdx + 1);
      const range = sheet.getRange(2, colIdx + 1, sheet.getLastRow() - 1, 1);
      
      // Filter out rules for this range
      rules = rules.filter(rule => {
        return !rule.getRanges().some(ruleRange => 
          ruleRange.getA1Notation() === range.getA1Notation());
      });
      
      // Create new rule for this column
      if (sheet.getLastRow() > 1) {  // Only add rules if there's data
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(
            `=AND(LEN(TRIM(${columnLetter}2))>0, NOT(REGEXMATCH(TRIM(${columnLetter}2), "^04\\d{2} \\d{3} \\d{3}$|^0[2378] \\d{4} \\d{4}$")))`
          )
          .setBackground("#FF0000")
          .setFontColor("#FFFFFF")
          .setRanges([range])
          .build();
        
        newRules.push(rule);
      }
    });

    // Apply all conditional formatting rules at once
    sheet.setConditionalFormatRules([...rules, ...newRules]);
    
    Performance.end('applyPhoneConditionalFormatting');
  }

  /*********************************
   * NAME PROCESSING
   *********************************/

  /**
   * Splits a full name into first and last name components.
   * @param {string} fullName The full name to split.
   * @returns {Object} Object with first and last name properties.
   */
  function _splitFullName(fullName) {
    if (!fullName || typeof fullName !== 'string') return { first: '', last: '' };
    
    const trimmed = fullName.trim();
    const lastSpace = trimmed.lastIndexOf(' ');
    
    if (lastSpace === -1) {
      return { first: trimmed, last: '' };
    }
    
    return {
      first: trimmed.substring(0, lastSpace),
      last: trimmed.substring(lastSpace + 1)
    };
  }

  /**
   * Processes parent names by splitting full names into first and last names.
   * Optimized to minimize API calls using batch operations.
   */
  function _processParentNames() {
    Performance.start('processParentNames');
    const { data, headers } = _getSheetData();
    
    // Get column indices for parent name fields
    const indices = {
      p1F: headers.indexOf(CONFIG.PARENT1_FNAME),
      p1L: headers.indexOf(CONFIG.PARENT1_LNAME),
      p2F: headers.indexOf(CONFIG.PARENT2_FNAME),
      p2L: headers.indexOf(CONFIG.PARENT2_LNAME)
    };

    // Validate that all required columns exist
    Object.entries(indices).forEach(([key, index]) => {
      if (index === -1) {
        const columnName = CONFIG[key.toUpperCase()] || key;
        throw new TransformError(`Missing column: ${columnName}`, 'MISSING_COLUMN');
      }
    });

    // Process names in batches to improve performance
    let modified = false;
    _processInBatches(data, batch => {
      batch.forEach(row => {
        ['p1', 'p2'].forEach(prefix => {
          const fIdx = indices[`${prefix}F`];
          const lIdx = indices[`${prefix}L`];
          const fullName = row[fIdx];
          const lastName = row[lIdx];
          
          if (fullName && (!lastName || lastName === '')) {
            const { first, last } = _splitFullName(fullName);
            row[fIdx] = first;
            row[lIdx] = last;
            modified = true;
          }
        });
      });
    });

    // Only update sheet if changes were made
    if (modified) {
      _updateSheetData(data);
    }
    
    Performance.end('processParentNames');
  }

  /**
   * Populates missing preferred names with first names.
   * Optimized to minimize API calls using batch operations.
   */
  function _processPreferredNames() {
    Performance.start('processPreferredNames');
    const { data, headers } = _getSheetData();
    
    const prefIdx = headers.indexOf(CONFIG.PREFERRED_NAME);
    const firstIdx = headers.indexOf(CONFIG.FIRST_NAME);
    
    if (prefIdx === -1 || firstIdx === -1) {
      throw new TransformError('Missing Preferred Name or First Name column', 'MISSING_COLUMN');
    }

    let modified = false;
    _processInBatches(data, batch => {
      batch.forEach(row => {
        if (!row[prefIdx] && row[firstIdx]) {
          row[prefIdx] = row[firstIdx];
          modified = true;
        }
      });
    });

    // Only update sheet if changes were made
    if (modified) {
      _updateSheetData(data);
    }
    
    Performance.end('processPreferredNames');
  }

  /*********************************
   * DUPLICATE VALIDATION
   *********************************/

  /**
   * Removes duplicate mobile numbers across member and parent fields.
   * Optimized to minimize API calls using batch operations.
   */
  function _deduplicateMobileNumbers() {
    Performance.start('deduplicateMobileNumbers');
    const { data, headers } = _getSheetData();
    const config = CONFIG;
    
    // Get column indices
    const memberMobileIdx = headers.indexOf(config.MEMBER_MOBILE);
    const parent1MobileIdx = headers.indexOf(config.PARENT1_MOBILE);
    const parent2MobileIdx = headers.indexOf(config.PARENT2_MOBILE);
    
    if ([memberMobileIdx, parent1MobileIdx, parent2MobileIdx].includes(-1)) {
      throw new TransformError('Missing mobile number columns', 'MISSING_COLUMN');
    }

    let modified = false;
    _processInBatches(data, batch => {
      batch.forEach(row => {
        const parentNumbers = new Set([
          _normalizePhoneNumber(row[parent1MobileIdx]),
          _normalizePhoneNumber(row[parent2MobileIdx])
        ].filter(n => n !== ''));
        
        const memberNumber = _normalizePhoneNumber(row[memberMobileIdx]);
        
        if (memberNumber && parentNumbers.has(memberNumber)) {
          row[memberMobileIdx] = '';
          modified = true;
        }
      });
    });

    // Only update sheet if changes were made
    if (modified) {
      _updateSheetData(data);
    }
    
    Performance.end('deduplicateMobileNumbers');
  }

  /**
   * Removes duplicate email addresses across member and parent fields.
   * Optimized to minimize API calls using batch operations.
   */
  function _deduplicateEmails() {
    Performance.start('deduplicateEmails');
    const { data, headers } = _getSheetData();
    const config = CONFIG;
    
    // Get column indices
    const emailIndices = {
      member: headers.indexOf(config.MEMBER_EMAIL),
      additional: headers.indexOf(config.ADDITIONAL_EMAILS),
      parent1: headers.indexOf(config.PARENT1_EMAIL),
      parent2: headers.indexOf(config.PARENT2_EMAIL)
    };
    
    if (Object.values(emailIndices).some(idx => idx === -1)) {
      throw new TransformError('Missing email columns', 'MISSING_COLUMN');
    }

    let modified = false;
    _processInBatches(data, batch => {
      batch.forEach(row => {
        const parentEmails = new Set([
          row[emailIndices.parent1]?.trim().toLowerCase(),
          row[emailIndices.parent2]?.trim().toLowerCase()
        ].filter(e => e));
        
        const memberEmail = row[emailIndices.member]?.trim().toLowerCase();
        
        // Remove member email if it duplicates parent email
        if (memberEmail && parentEmails.has(memberEmail)) {
          row[emailIndices.member] = '';
          modified = true;
        }
        
        // Remove duplicate emails from additional emails field
        if (row[emailIndices.additional]) {
          const additionalEmails = row[emailIndices.additional]
            .split(',')
            .map(e => e.trim().toLowerCase())
            .filter(e => e && !parentEmails.has(e) && e !== memberEmail);
          
          // Deduplicate additional emails
          const uniqueEmails = [...new Set(additionalEmails)].join(', ');
          
          if (row[emailIndices.additional] !== uniqueEmails) {
            row[emailIndices.additional] = uniqueEmails;
            modified = true;
          }
        }
      });
    });

    // Only update sheet if changes were made
    if (modified) {
      _updateSheetData(data);
    }
    
    Performance.end('deduplicateEmails');
  }

  /*********************************
   * BATCH PROCESSING FOR LARGE DATASETS
   *********************************/

  /**
   * Processes data in batches to prevent hitting execution time limits.
   * @param {Array<Array>} data The data to process.
   * @param {Function} processor Function to process each batch.
   * @param {number} [batchSize=500] Size of each batch.
   */
  function _processInBatches(data, processor, batchSize = 500) {
    if (!data || !data.length) return;
    
    for (let i = 0; i < data.length; i += batchSize) {
      const batch = data.slice(i, i + batchSize);
      processor(batch);
      
      // Add small delay between batches for long datasets to prevent timeouts
      if (i + batchSize < data.length) {
        Utilities.sleep(50);
      }
    }
  }

  /**
   * Sorts data by membership number.
   * Optimized to handle large datasets.
   */
  function _sortByMemberNumber() {
    Performance.start('sortByMemberNumber');
    const sheet = _getSheet();
    
    // Get the index of the membership number column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const memberNumIdx = headers.indexOf(CONFIG.MEMBER_NUMBER);
    
    if (memberNumIdx === -1) {
      throw new TransformError('Membership Number column not found', 'MISSING_COLUMN');
    }

    // Only sort if there's data to sort
    if (sheet.getLastRow() > 1) {
      // Perform the sort in a single operation
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
        .sort({ column: memberNumIdx + 1, ascending: true });
      
      // Clear cache since data order changed
      CacheService.getScriptCache().remove('sheetData');
    }
    
    Performance.end('sortByMemberNumber');
  }

  /**
   * Runs all transform operations in sequence.
   * Uses performance monitoring and provides user feedback.
   */
  function runAllTransforms() {
    const ui = SpreadsheetApp.getUi();
    
    try {
      Performance.start('runAllTransforms');
      
      // Execute all transforms in sequence
      _processPhoneColumns();
      _deduplicateMobileNumbers();
      _processParentNames();
      _processPreferredNames();
      _deduplicateEmails();
      _sortByMemberNumber();
      
      const totalTime = Performance.end('runAllTransforms');
      
      ui.alert(
        'All transforms completed successfully', 
        `Operation completed in ${totalTime/1000} seconds.\n\n` +
        'Check highlighted invalid numbers and duplicates.',
        ui.ButtonSet.OK
      );
    } catch (e) {
      _handleError('Batch processing failed', e);
    }
  }

  /*********************************
   * ERROR HANDLING
   *********************************/

  /**
   * Custom error class for transform-specific errors.
   */
  class TransformError extends Error {
    /**
     * @param {string} message Human-readable error message
     * @param {string} code Error code for programmatic handling
     */
    constructor(message, code) {
      super(message);
      this.code = code;
      this.name = 'TransformError';
    }
  }

  /**
   * Handles errors with user-friendly alerts and logging.
   * @param {string} message User-friendly message prefix
   * @param {Error} error The error object
   */
  function _handleError(message, error) {
    const ui = SpreadsheetApp.getUi();
    
    // Log detailed error information for debugging
    Logger.log(`[${new Date().toISOString()}] ERROR: ${message}\n${error.stack}`);
    
    // Prepare user-friendly error message
    const errorMessage = error instanceof TransformError
      ? `${error.message} (Code: ${error.code})`
      : `System Error: ${error.message}`;
    
    // Show alert to user
    ui.alert(
      'Error',
      `${message}\n\n${errorMessage}`,
      ui.ButtonSet.OK
    );
  }

  /*********************************
   * PUBLIC API
   *********************************/

  return {
    /**
     * Normalizes phone numbers across all configured phone columns.
     */
    normalizePhoneNumbers: function() {
      try {
        _processPhoneColumns();
        SpreadsheetApp.getUi().alert('Phone numbers processed\nInvalid numbers highlighted in red');
      } catch (e) {
        _handleError('Phone processing failed', e);
      }
    },
    
    /**
     * Processes parent names by splitting full names into first and last names.
     */
    processParentNames: function() {
      try {
        _processParentNames();
        SpreadsheetApp.getUi().alert('Parent names processed');
      } catch (e) {
        _handleError('Name processing failed', e);
      }
    },
    
    /**
     * Populates missing preferred names with first names.
     */
    processPreferredNames: function() {
      try {
        _processPreferredNames();
        SpreadsheetApp.getUi().alert('Preferred names updated');
      } catch (e) {
        _handleError('Preferred name update failed', e);
      }
    },
    
    /**
     * Sorts data by membership number.
     */
    sortByMemberNumber: function() {
      try {
        _sortByMemberNumber();
        SpreadsheetApp.getUi().alert('Sorted by Member Number');
      } catch (e) {
        _handleError('Sorting failed', e);
      }
    },
    
    /**
     * Runs all transform operations.
     */
    runAllTransforms: runAllTransforms,
    
    /**
     * Updates a configuration value.
     * @param {string} key Configuration key
     * @param {*} value New configuration value
     */
    updateConfig: updateConfig,
    
    /**
     * Tests phone number normalization with sample inputs.
     */
    testPhoneNormalization: function() {
      const tests = {
        '0412345678': '0412 345 678',
        '+61412345678': '0412 345 678',
        '61412345678': '0412 345 678',
        '412345678': '0412 345 678',
        '0298765432': '02 9876 5432',
        '298765432': '02 9876 5432',
        '98765432': '02 9876 5432',
        '': '',
        'invalid!': '',
        undefined: ''
      };
      
      let results = 'Phone Normalization Test Results:\n';
      let passed = 0;
      
      Object.entries(tests).forEach(([input, expected]) => {
        const result = _normalizePhoneNumber(input);
        const status = result === expected ? 'PASSED' : 'FAILED';
        if (status === 'PASSED') passed++;
        
        results += `${status}: "${input}" → "${result}" (Expected: "${expected}")\n`;
        Logger.log(`Test ${status}: "${input}" → "${result}" (Expected: "${expected}")`);
      });
      
      SpreadsheetApp.getUi().alert(
        `Phone Tests: ${passed}/${Object.keys(tests).length} passed`,
        results,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    },
    
    /**
     * Removes duplicate mobile numbers.
     */
    deduplicateMobileNumbers: function() {
      try {
        _deduplicateMobileNumbers();
        SpreadsheetApp.getUi().alert('Duplicate mobile numbers removed from member field.');
      } catch (e) {
        _handleError('Mobile deduplication failed', e);
      }
    },
    
    /**
     * Removes duplicate email addresses.
     */
    deduplicateEmails: function() {
      try {
        _deduplicateEmails();
        SpreadsheetApp.getUi().alert('Duplicate emails removed from member and additional fields.');
      } catch (e) {
        _handleError('Email deduplication failed', e);
      }
    },
    
    /**
     * Resets all configuration to default values.
     */
    resetAllConfigs: resetAllConfigs,
    
    /**
     * Clears all caches to force fresh data.
     */
    clearCaches: function() {
      CacheService.getScriptCache().removeAll([
        'TRANSFORM_CONFIG', 'targetSheetName', 'headerIndices', 'sheetData'
      ]);
      SpreadsheetApp.getUi().alert('All caches cleared.');
    }
  };
})();

/* Public function aliases for menu integration */
function normalizePhoneNumbers() { TransformsModule.normalizePhoneNumbers(); }
function applyParentNameSplitting() { TransformsModule.processParentNames(); }
function applyPreferredNamePopulation() { TransformsModule.processPreferredNames(); }
function sortByMemberNumber() { TransformsModule.sortByMemberNumber(); }
function runAllTransforms() { TransformsModule.runAllTransforms(); }
function testPhoneNormalization() { TransformsModule.testPhoneNormalization(); }
function updateConfig(key, value) { TransformsModule.updateConfig(key, value); }
function deduplicateMobileNumbers() { TransformsModule.deduplicateMobileNumbers(); }
function deduplicateEmails() { TransformsModule.deduplicateEmails(); }
function resetAllConfigsMenu() { TransformsModule.resetAllConfigs(); }
function clearTransformCaches() { TransformsModule.clearCaches(); }
