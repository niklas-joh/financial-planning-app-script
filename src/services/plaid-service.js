/**
 * @fileoverview Plaid Service Module for Financial Planning Tools.
 * Handles API communication with Plaid for bank account linking and transaction retrieval.
 * @module services/plaid-service
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Plaid Service - Handles Plaid API integration for bank connections and transactions.
 * @namespace FinancialPlanner.PlaidService
 */
FinancialPlanner.PlaidService = (function() {
  /**
   * Gets the appropriate Plaid API URL based on environment setting.
   * @private
   * @returns {string} The Plaid API base URL.
   */
  function getApiUrl() {
    const plaidConfig = FinancialPlanner.Config.getSection('PLAID');
    return plaidConfig.API_URL || 'https://sandbox.plaid.com';
  }

  /**
   * Gets API credentials from Script Properties.
   * @private
   * @returns {{clientId: string, secret: string}} The Plaid credentials.
   * @throws {Error} If credentials are not configured.
   */
  function getCredentials() {
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty('PLAID_CLIENT_ID');
    const secret = props.getProperty('PLAID_SECRET');
    
    if (!clientId || !secret) {
      throw FinancialPlanner.ErrorService.create(
        'Plaid credentials not configured. Please set PLAID_CLIENT_ID and PLAID_SECRET in Script Properties.',
        { severity: 'high' }
      );
    }
    
    return { clientId: clientId, secret: secret };
  }

  /**
   * Maps a Plaid category to the application's transaction type.
   * @private
   * @param {Array<string>} plaidCategories - Plaid category array.
   * @returns {string} The mapped transaction type.
   */
  function mapCategory(plaidCategories) {
    if (!plaidCategories || plaidCategories.length === 0) {
      return 'Extra';
    }
    
    const categoryMap = FinancialPlanner.Config.getSection('PLAID').CATEGORY_MAP || {};
    const primaryCategory = plaidCategories[0];
    
    return categoryMap[primaryCategory] || 'Extra';
  }

  /**
   * Gets the stored cursor for transaction sync.
   * @private
   * @returns {string|null} The stored cursor or null.
   */
  function getCursor() {
    return PropertiesService.getScriptProperties().getProperty('PLAID_SYNC_CURSOR');
  }

  /**
   * Saves the cursor for next transaction sync.
   * @private
   * @param {string} cursor - The cursor to save.
   */
  function saveCursor(cursor) {
    PropertiesService.getScriptProperties().setProperty('PLAID_SYNC_CURSOR', cursor);
  }

  /**
   * Helper function to safely convert value, handling null/undefined.
   * @private
   * @param {*} value - The value to convert.
   * @param {*} defaultValue - The default value if null/undefined.
   * @returns {*} The safe value.
   */
  function safeValue(value, defaultValue) {
    return (value !== null && value !== undefined) ? value : defaultValue;
  }

  /**
   * Helper function to stringify objects/arrays.
   * @private
   * @param {*} value - The value to stringify.
   * @returns {string} The stringified value.
   */
  function stringify(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') return JSON.stringify(value);
    return String(value);
  }

  /**
   * Flattens a Plaid transaction object, expanding nested objects into separate fields.
   * @private
   * @param {object} tx - The transaction object to flatten.
   * @returns {object} The flattened transaction object.
   */
  function flattenTransaction(tx) {
    const flat = { ...tx };
    
    // Handle location nested object
    if (tx.location) {
      flat['location.address'] = tx.location.address;
      flat['location.city'] = tx.location.city;
      flat['location.region'] = tx.location.region;
      flat['location.postal_code'] = tx.location.postal_code;
      flat['location.country'] = tx.location.country;
      flat['location.lat'] = tx.location.lat;
      flat['location.lon'] = tx.location.lon;
      flat['location.store_number'] = tx.location.store_number;
      delete flat.location;
    }
    
    // Handle payment_meta nested object
    if (tx.payment_meta) {
      flat['payment_meta.reference_number'] = tx.payment_meta.reference_number;
      flat['payment_meta.ppd_id'] = tx.payment_meta.ppd_id;
      flat['payment_meta.payee'] = tx.payment_meta.payee;
      flat['payment_meta.by_order_of'] = tx.payment_meta.by_order_of;
      flat['payment_meta.payer'] = tx.payment_meta.payer;
      flat['payment_meta.payment_method'] = tx.payment_meta.payment_method;
      flat['payment_meta.payment_processor'] = tx.payment_meta.payment_processor;
      flat['payment_meta.reason'] = tx.payment_meta.reason;
      delete flat.payment_meta;
    }
    
    // Handle personal_finance_category nested object
    if (tx.personal_finance_category) {
      flat['personal_finance_category.primary'] = tx.personal_finance_category.primary;
      flat['personal_finance_category.detailed'] = tx.personal_finance_category.detailed;
      flat['personal_finance_category.confidence_level'] = tx.personal_finance_category.confidence_level;
      delete flat.personal_finance_category;
    }
    
    // Handle arrays - convert to comma-separated strings
    if (Array.isArray(flat.category)) {
      flat.category = flat.category.join(', ');
    }
    if (Array.isArray(flat.counterparties)) {
      flat.counterparties = stringify(flat.counterparties);
    }
    
    return flat;
  }

  // Public API
  return {
    /**
     * Creates a Plaid Link token for initiating the bank connection flow.
     * @returns {{link_token: string, expiration: string}} The link token response.
     * @memberof FinancialPlanner.PlaidService
     */
    createLinkToken: function() {
      const url = getApiUrl() + '/link/token/create';
      const credentials = getCredentials();
      
      const payload = {
        client_id: credentials.clientId,
        secret: credentials.secret,
        user: {
          client_user_id: SpreadsheetApp.getActiveSpreadsheet().getId()
        },
        client_name: 'Financial Planning Tools',
        products: ['transactions'],
        country_codes: ['US'],
        language: 'en'
      };
      
      try {
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
            'Failed to create Plaid Link token',
            { responseCode: responseCode, response: responseText, severity: 'high' }
          );
        }
        
        return JSON.parse(responseText);
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to create Plaid Link token');
        throw error;
      }
    },

    /**
     * Exchanges a public token for an access token.
     * @param {string} publicToken - The public token from Plaid Link.
     * @returns {{access_token: string, item_id: string}} The access token response.
     * @memberof FinancialPlanner.PlaidService
     */
    exchangePublicToken: function(publicToken) {
      const url = getApiUrl() + '/item/public_token/exchange';
      const credentials = getCredentials();
      
      const payload = {
        client_id: credentials.clientId,
        secret: credentials.secret,
        public_token: publicToken
      };
      
      try {
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
            'Failed to exchange public token',
            { responseCode: responseCode, response: responseText, severity: 'high' }
          );
        }
        
        const result = JSON.parse(responseText);
        
        // Store access token in Script Properties
        PropertiesService.getScriptProperties().setProperty('PLAID_ACCESS_TOKEN', result.access_token);
        
        return result;
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to exchange public token');
        throw error;
      }
    },

    /**
     * Syncs transactions from Plaid using the /transactions/sync endpoint.
     * Uses cursor-based pagination to fetch all available transaction updates.
     * @returns {{added: Array<object>, modified: Array<object>, removed: Array<object>}} Sync results with added, modified, and removed transactions.
     * @memberof FinancialPlanner.PlaidService
     */
    syncTransactions: function() {
      const url = getApiUrl() + '/transactions/sync';
      const credentials = getCredentials();
      const accessToken = PropertiesService.getScriptProperties().getProperty('PLAID_ACCESS_TOKEN');
      
      if (!accessToken) {
        throw FinancialPlanner.ErrorService.create(
          'No bank account connected. Please connect your bank account first.',
          { severity: 'medium' }
        );
      }
      
      let allAdded = [];
      let allModified = [];
      let allRemoved = [];
      let nextCursor;
      let hasMore = true;
      let currentCursor = getCursor(); // null for first sync = full history
      
      try {
        Logger.log('Starting transaction sync' + (currentCursor ? ' with cursor' : ' (full history)'));
        
        // Pagination loop to fetch all pages
        while (hasMore) {
          const payload = {
            client_id: credentials.clientId,
            secret: credentials.secret,
            access_token: accessToken,
            count: 500 // Maximum transactions per request
          };
          
          // Add cursor if available
          if (currentCursor) {
            payload.cursor = currentCursor;
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
              'Failed to sync transactions from Plaid',
              { responseCode: responseCode, response: responseText, severity: 'high' }
            );
          }
          
          const data = JSON.parse(responseText);
          
          // Accumulate transaction updates
          allAdded = allAdded.concat(data.added);
          allModified = allModified.concat(data.modified);
          allRemoved = allRemoved.concat(data.removed);
          nextCursor = data.next_cursor;
          hasMore = data.has_more;
          
          Logger.log('Fetched page: ' + data.added.length + ' added, ' + 
                     data.modified.length + ' modified, ' + 
                     data.removed.length + ' removed. Has more: ' + hasMore);
          
          // Update cursor for next iteration
          if (hasMore) {
            currentCursor = nextCursor;
          }
        }
        
        // Save cursor for next sync
        saveCursor(nextCursor);
        Logger.log('Sync complete. Total: ' + allAdded.length + ' added, ' + 
                   allModified.length + ' modified, ' + 
                   allRemoved.length + ' removed');
        
        return {
          added: allAdded,
          modified: allModified,
          removed: allRemoved
        };
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to sync transactions from Plaid');
        throw error;
      }
    },

    /**
     * Imports Plaid transaction sync results to the Transactions sheet.
     * Handles added, modified, and removed transactions with dynamic column creation.
     * @param {object} syncResults - Sync results with added, modified, and removed arrays.
     * @returns {number} Number of transactions processed.
     * @memberof FinancialPlanner.PlaidService
     */
    importToSheet: function(syncResults) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetNames = FinancialPlanner.Config.getSheetNames();
      let transactionSheet = ss.getSheetByName(sheetNames.TRANSACTIONS);
      
      // Combine all transactions to check if there's anything to process
      const allTransactions = [].concat(syncResults.added || [], syncResults.modified || []);
      
      if (allTransactions.length === 0 && (!syncResults.removed || syncResults.removed.length === 0)) {
        Logger.log('No transactions to process');
        return 0;
      }
      
      Logger.log('Processing sync results: ' + 
                 (syncResults.added ? syncResults.added.length : 0) + ' added, ' +
                 (syncResults.modified ? syncResults.modified.length : 0) + ' modified, ' +
                 (syncResults.removed ? syncResults.removed.length : 0) + ' removed');
      
      // Get or create headers from first transaction
      let headers;
      if (!transactionSheet || transactionSheet.getLastRow() === 0) {
        // New sheet - create headers dynamically from first transaction
        if (allTransactions.length > 0) {
          const firstTx = allTransactions[0];
          const flattened = flattenTransaction(firstTx);
          headers = Object.keys(flattened);
          headers.push('deleted'); // Add deleted flag column
          
          Logger.log('Creating new sheet with ' + headers.length + ' dynamic columns');
          
          if (!transactionSheet) {
            transactionSheet = ss.insertSheet(sheetNames.TRANSACTIONS);
          }
          transactionSheet.getRange(1, 1, 1, headers.length)
            .setValues([headers])
            .setFontWeight('bold');
        } else {
          Logger.log('No transactions to create headers from');
          return 0;
        }
      } else {
        // Existing sheet - use existing headers
        headers = transactionSheet.getRange(1, 1, 1, transactionSheet.getLastColumn()).getValues()[0];
        Logger.log('Using existing headers: ' + headers.length + ' columns');
      }
      
      const txIdIndex = headers.indexOf('transaction_id');
      
      if (txIdIndex === -1) {
        throw FinancialPlanner.ErrorService.create(
          'transaction_id column not found in sheet headers',
          { severity: 'high' }
        );
      }
      
      // Process added transactions (simple append)
      if (syncResults.added && syncResults.added.length > 0) {
        Logger.log('Adding ' + syncResults.added.length + ' new transactions');
        
        const addedRows = syncResults.added.map(function(tx) {
          const flattened = flattenTransaction(tx);
          return headers.map(function(header) {
            if (header === 'deleted') {
              return false;
            }
            const value = flattened[header];
            // Parse date fields
            if ((header.includes('date') || header.includes('datetime')) && value) {
              return new Date(value);
            }
            return safeValue(value, '');
          });
        });
        
        const lastRow = transactionSheet.getLastRow();
        transactionSheet.getRange(lastRow + 1, 1, addedRows.length, headers.length)
          .setValues(addedRows);
        Logger.log('Added transactions at row ' + (lastRow + 1));
      }
      
      // Process modified and removed transactions (single sheet scan for efficiency)
      if ((syncResults.modified && syncResults.modified.length > 0) || 
          (syncResults.removed && syncResults.removed.length > 0)) {
        
        Logger.log('Processing ' + 
                   (syncResults.modified ? syncResults.modified.length : 0) + ' modified and ' +
                   (syncResults.removed ? syncResults.removed.length : 0) + ' removed transactions');
        
        const data = transactionSheet.getDataRange().getValues();
        const modifiedMap = {};
        const removedSet = new Set();
        
        // Build lookup maps
        if (syncResults.modified) {
          syncResults.modified.forEach(function(tx) {
            modifiedMap[tx.transaction_id] = tx;
          });
        }
        
        if (syncResults.removed) {
          syncResults.removed.forEach(function(tx) {
            removedSet.add(tx.transaction_id);
          });
        }
        
        // Single pass through sheet to update rows
        for (let i = 1; i < data.length; i++) {
          const rowTxId = data[i][txIdIndex];
          
          // Update modified transactions
          if (modifiedMap[rowTxId]) {
            const tx = modifiedMap[rowTxId];
            const flattened = flattenTransaction(tx);
            const rowData = headers.map(function(header) {
              if (header === 'deleted') {
                return false; // Reset deleted flag for modified transactions
              }
              const value = flattened[header];
              if ((header.includes('date') || header.includes('datetime')) && value) {
                return new Date(value);
              }
              return safeValue(value, '');
            });
            transactionSheet.getRange(i + 1, 1, 1, headers.length).setValues([rowData]);
          }
          
          // Mark removed transactions
          if (removedSet.has(rowTxId)) {
            const deletedColIndex = headers.indexOf('deleted') + 1;
            if (deletedColIndex > 0) {
              transactionSheet.getRange(i + 1, deletedColIndex).setValue(true);
            }
          }
        }
        
        Logger.log('Updated existing transactions');
      }
      
      const totalProcessed = (syncResults.added ? syncResults.added.length : 0) +
                             (syncResults.modified ? syncResults.modified.length : 0) +
                             (syncResults.removed ? syncResults.removed.length : 0);
      
      Logger.log('Successfully processed ' + totalProcessed + ' transactions');
      return totalProcessed;
    }
  };
})();
