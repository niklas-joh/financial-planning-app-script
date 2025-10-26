/**
 * @fileoverview Plaid Institutions Module - Institution fetching and sheet operations.
 * Handles fetching institutions from Plaid and importing them to Google Sheets.
 * @module services/plaid-institutions
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Plaid Institutions - Handles institution fetching and import operations.
 * @namespace FinancialPlanner.PlaidInstitutions
 */
FinancialPlanner.PlaidInstitutions = (function() {
  /**
   * Fetches a single page of institutions from Plaid.
   * @private
   * @param {number} count - Number of institutions to fetch (1-500).
   * @param {number} offset - Starting position.
   * @param {string} countryCode - ISO-3166-1 alpha-2 country code.
   * @param {object} options - Optional metadata flags.
   * @returns {Array<object>} Array of institution objects.
   */
  function fetchPage(count, offset, countryCode, options) {
    const url = FinancialPlanner.PlaidClient.getApiUrl() + '/institutions/get';
    const credentials = FinancialPlanner.PlaidClient.getCredentials();
    
    const payload = {
      client_id: credentials.clientId,
      secret: credentials.secret,
      count: count,
      offset: offset,
      country_codes: [countryCode]
    };
    
    // Add optional metadata flags if provided
    if (options) {
      if (options.include_optional_metadata) {
        payload.options = payload.options || {};
        payload.options.include_optional_metadata = true;
      }
      if (options.include_auth_metadata) {
        payload.options = payload.options || {};
        payload.options.include_auth_metadata = true;
      }
      if (options.include_payment_initiation_metadata) {
        payload.options = payload.options || {};
        payload.options.include_payment_initiation_metadata = true;
      }
    }
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      throw FinancialPlanner.ErrorService.create(
        'Failed to fetch institutions from Plaid',
        { 
          responseCode: responseCode, 
          response: responseText, 
          countryCode: countryCode,
          offset: offset,
          severity: 'high' 
        }
      );
    }
    
    const data = JSON.parse(responseText);
    return data.institutions || [];
  }

  // Public API
  return {
    /**
     * Fetches ALL institutions from Plaid for specified country codes with pagination.
     * @param {Array<string>} countryCodes - Array of ISO-3166-1 alpha-2 country codes (e.g., ['DE', 'FR', 'BE', 'NL']).
     * @param {object} [options] - Optional metadata flags.
     * @param {boolean} [options.include_optional_metadata] - Include logos, URLs, colors.
     * @param {boolean} [options.include_auth_metadata] - Include auth method support.
     * @param {boolean} [options.include_payment_initiation_metadata] - Include payment configurations.
     * @returns {Array<object>} Array of all institutions from all countries.
     * @memberof FinancialPlanner.PlaidInstitutions
     */
    fetchAll: function(countryCodes, options) {
      let allInstitutions = [];
      
      Logger.log('Starting institutions fetch for countries: ' + countryCodes.join(', '));
      
      try {
        // Loop through each country
        for (let c = 0; c < countryCodes.length; c++) {
          const countryCode = countryCodes[c];
          let offset = 0;
          let hasMore = true;
          const pageSize = 500; // Maximum allowed by Plaid
          
          Logger.log('Fetching institutions for ' + countryCode + '...');
          
          // Paginate through all institutions for this country
          while (hasMore) {
            const institutions = fetchPage(pageSize, offset, countryCode, options);
            
            if (institutions.length === 0) {
              hasMore = false;
            } else {
              allInstitutions = allInstitutions.concat(institutions);
              offset += institutions.length;
              
              Logger.log('Fetched ' + institutions.length + ' institutions for ' + countryCode + 
                         ' (total for country: ' + offset + ')');
              
              // If we got fewer than pageSize, we've reached the end
              if (institutions.length < pageSize) {
                hasMore = false;
              }
            }
          }
          
          Logger.log('Completed ' + countryCode + ': ' + offset + ' institutions');
        }
        
        Logger.log('Fetch complete. Total institutions: ' + allInstitutions.length);
        return allInstitutions;
        
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to fetch institutions from Plaid');
        throw error;
      }
    },

    /**
     * Imports institutions to the Institutions sheet with dynamic column creation.
     * @param {Array<object>} institutions - Array of institution objects from Plaid.
     * @returns {number} Number of institutions imported.
     * @memberof FinancialPlanner.PlaidInstitutions
     */
    importToSheet: function(institutions) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetNames = FinancialPlanner.Config.getSheetNames();
      let institutionSheet = ss.getSheetByName(sheetNames.INSTITUTIONS);
      
      if (!institutions || institutions.length === 0) {
        Logger.log('No institutions to import');
        return 0;
      }
      
      Logger.log('Importing ' + institutions.length + ' institutions to sheet');
      
      // Flatten the first institution to get all possible headers
      const firstInstitution = institutions[0];
      const flattened = FinancialPlanner.PlaidClient.flattenObject(firstInstitution);
      const headers = Object.keys(flattened);
      
      Logger.log('Creating sheet with ' + headers.length + ' dynamic columns');
      
      // Create or clear the sheet
      if (!institutionSheet) {
        institutionSheet = ss.insertSheet(sheetNames.INSTITUTIONS);
      } else {
        institutionSheet.clear();
      }
      
      // Write headers
      institutionSheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      // Prepare data rows
      const rows = institutions.map(function(institution) {
        const flattened = FinancialPlanner.PlaidClient.flattenObject(institution);
        return headers.map(function(header) {
          const value = flattened[header];
          // Parse date fields
          if ((header.includes('date') || header.includes('datetime') || header.includes('_change')) && value && typeof value === 'string') {
            try {
              return new Date(value);
            } catch (e) {
              return value;
            }
          }
          return FinancialPlanner.PlaidClient.safeValue(value, '');
        });
      });
      
      // Write all data at once for efficiency
      if (rows.length > 0) {
        institutionSheet.getRange(2, 1, rows.length, headers.length)
          .setValues(rows);
      }
      
      // Auto-resize first few columns for readability
      const columnsToResize = Math.min(5, headers.length);
      for (let i = 1; i <= columnsToResize; i++) {
        institutionSheet.autoResizeColumn(i);
      }
      
      // Freeze header row
      institutionSheet.setFrozenRows(1);
      
      Logger.log('Successfully imported ' + institutions.length + ' institutions');
      return institutions.length;
    }
  };
})();
