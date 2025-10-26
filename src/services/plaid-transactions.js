/**
 * @fileoverview Plaid Transactions Module - Transaction sync and sheet operations.
 * Handles fetching transactions from Plaid and importing them to Google Sheets.
 * @module services/plaid-transactions
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Plaid Transactions - Handles transaction sync and import operations.
 * @namespace FinancialPlanner.PlaidTransactions
 */
FinancialPlanner.PlaidTransactions = (function() {
  /**
   * Gets the stored cursor for transaction sync for the current environment.
   * @private
   * @returns {string|null} The stored cursor or null.
   */
  function getCursor() {
    const env = FinancialPlanner.SettingsService.getPlaidEnvironment();
    const key = 'PLAID_' + env.toUpperCase() + '_SYNC_CURSOR';
    return PropertiesService.getScriptProperties().getProperty(key);
  }

  /**
   * Saves the cursor for next transaction sync for the current environment.
   * @private
   * @param {string} cursor - The cursor to save.
   */
  function saveCursor(cursor) {
    const env = FinancialPlanner.SettingsService.getPlaidEnvironment();
    const key = 'PLAID_' + env.toUpperCase() + '_SYNC_CURSOR';
    PropertiesService.getScriptProperties().setProperty(key, cursor);
  }

  /**
   * Resets the cursor to force a full sync on next call for the current environment.
   * @private
   */
  function resetCursor() {
    const env = FinancialPlanner.SettingsService.getPlaidEnvironment();
    const key = 'PLAID_' + env.toUpperCase() + '_SYNC_CURSOR';
    PropertiesService.getScriptProperties().deleteProperty(key);
  }

  // Public API
  return {
    /**
     * Syncs transactions from Plaid using the /transactions/sync endpoint.
     * Uses cursor-based pagination to fetch all available transaction updates.
     * @returns {{added: Array<object>, modified: Array<object>, removed: Array<object>}} Sync results with added, modified, and removed transactions.
     * @memberof FinancialPlanner.PlaidTransactions
     */
    syncAll: function() {
      const url = FinancialPlanner.PlaidClient.getApiUrl() + '/transactions/sync';
      const credentials = FinancialPlanner.PlaidClient.getCredentials();
      
      const env = FinancialPlanner.SettingsService.getPlaidEnvironment();
      const tokenKey = 'PLAID_' + env.toUpperCase() + '_ACCESS_TOKEN';
      const accessToken = PropertiesService.getScriptProperties().getProperty(tokenKey);
      
      if (!accessToken) {
        throw FinancialPlanner.ErrorService.create(
          'No bank account connected for ' + env + ' environment. Please connect your bank account first.',
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
     * Resets the sync cursor and fetches all transactions from scratch.
     * Useful for development, testing, or re-syncing all data.
     * @returns {{added: Array<object>, modified: Array<object>, removed: Array<object>}} Sync results with all transactions.
     * @memberof FinancialPlanner.PlaidTransactions
     */
    resetAndSyncAll: function() {
      Logger.log('Resetting cursor and fetching all transactions...');
      resetCursor();
      return this.syncAll();
    },

    /**
     * Imports Plaid transaction sync results to the Transactions sheet.
     * Handles added, modified, and removed transactions with dynamic column creation.
     * @param {object} syncResults - Sync results with added, modified, and removed arrays.
     * @returns {number} Number of transactions processed.
     * @memberof FinancialPlanner.PlaidTransactions
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
          const flattened = FinancialPlanner.PlaidClient.flattenObject(firstTx);
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
          const flattened = FinancialPlanner.PlaidClient.flattenObject(tx);
          return headers.map(function(header) {
            if (header === 'deleted') {
              return false;
            }
            const value = flattened[header];
            // Parse date fields
            if ((header.includes('date') || header.includes('datetime')) && value) {
              return new Date(value);
            }
            return FinancialPlanner.PlaidClient.safeValue(value, '');
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
            const flattened = FinancialPlanner.PlaidClient.flattenObject(tx);
            const rowData = headers.map(function(header) {
              if (header === 'deleted') {
                return false; // Reset deleted flag for modified transactions
              }
              const value = flattened[header];
              if ((header.includes('date') || header.includes('datetime')) && value) {
                return new Date(value);
              }
              return FinancialPlanner.PlaidClient.safeValue(value, '');
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
