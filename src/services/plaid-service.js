/**
 * @fileoverview Plaid Service Module for Financial Planning Tools.
 * Handles API communication with Plaid for bank account linking and transaction retrieval.
 * This module is designed to be instantiated by 00_module_loader.js.
 * @module services/plaid-service
 */

/**
 * IIFE to encapsulate the PlaidServiceModule logic.
 * @returns {function} The PlaidServiceModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const PlaidServiceModule = (function() {
  /**
   * Constructor for the PlaidServiceModule.
   * @param {ConfigModule} configInstance - An instance of ConfigModule.
   * @param {ErrorServiceModule} errorServiceInstance - An instance of ErrorServiceModule.
   * @constructor
   * @alias PlaidServiceModule
   * @memberof module:services/plaid-service
   */
  function PlaidServiceModuleConstructor(configInstance, errorServiceInstance) {
    this.config = configInstance;
    this.errorService = errorServiceInstance;
  }

  /**
   * Gets the appropriate Plaid API URL based on environment setting.
   * @private
   * @returns {string} The Plaid API base URL.
   */
  PlaidServiceModuleConstructor.prototype._getApiUrl = function() {
    const plaidConfig = this.config.getSection('PLAID');
    return plaidConfig.API_URL || 'https://sandbox.plaid.com';
  };

  /**
   * Gets API credentials from Script Properties.
   * @private
   * @returns {{clientId: string, secret: string}} The Plaid credentials.
   * @throws {Error} If credentials are not configured.
   */
  PlaidServiceModuleConstructor.prototype._getCredentials = function() {
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty('PLAID_CLIENT_ID');
    const secret = props.getProperty('PLAID_SECRET');
    
    if (!clientId || !secret) {
      throw this.errorService.create(
        'Plaid credentials not configured. Please set PLAID_CLIENT_ID and PLAID_SECRET in Script Properties.',
        { severity: 'high' }
      );
    }
    
    return { clientId: clientId, secret: secret };
  };

  /**
   * Creates a Plaid Link token for initiating the bank connection flow.
   * @returns {{link_token: string, expiration: string}} The link token response.
   * @memberof PlaidServiceModule
   */
  PlaidServiceModuleConstructor.prototype.createLinkToken = function() {
    const url = this._getApiUrl() + '/link/token/create';
    const credentials = this._getCredentials();
    
    const payload = {
      client_id: credentials.clientId,
      secret: credentials.secret,
      user: {
        client_user_id: Session.getActiveUser().getEmail()
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
        throw this.errorService.create(
          'Failed to create Plaid Link token',
          { responseCode: responseCode, response: responseText, severity: 'high' }
        );
      }
      
      return JSON.parse(responseText);
    } catch (error) {
      this.errorService.handle(error, 'Failed to create Plaid Link token');
      throw error;
    }
  };

  /**
   * Exchanges a public token for an access token.
   * @param {string} publicToken - The public token from Plaid Link.
   * @returns {{access_token: string, item_id: string}} The access token response.
   * @memberof PlaidServiceModule
   */
  PlaidServiceModuleConstructor.prototype.exchangePublicToken = function(publicToken) {
    const url = this._getApiUrl() + '/item/public_token/exchange';
    const credentials = this._getCredentials();
    
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
        throw this.errorService.create(
          'Failed to exchange public token',
          { responseCode: responseCode, response: responseText, severity: 'high' }
        );
      }
      
      const result = JSON.parse(responseText);
      
      // Store access token in Script Properties
      PropertiesService.getScriptProperties().setProperty('PLAID_ACCESS_TOKEN', result.access_token);
      
      return result;
    } catch (error) {
      this.errorService.handle(error, 'Failed to exchange public token');
      throw error;
    }
  };

  /**
   * Retrieves transactions from Plaid for a specified date range.
   * @param {string} startDate - Start date in YYYY-MM-DD format.
   * @param {string} endDate - End date in YYYY-MM-DD format.
   * @returns {{transactions: Array<object>}} The transactions response.
   * @memberof PlaidServiceModule
   */
  PlaidServiceModuleConstructor.prototype.getTransactions = function(startDate, endDate) {
    const url = this._getApiUrl() + '/transactions/get';
    const credentials = this._getCredentials();
    const accessToken = PropertiesService.getScriptProperties().getProperty('PLAID_ACCESS_TOKEN');
    
    if (!accessToken) {
      throw this.errorService.create(
        'No bank account connected. Please connect your bank account first.',
        { severity: 'medium' }
      );
    }
    
    const payload = {
      client_id: credentials.clientId,
      secret: credentials.secret,
      access_token: accessToken,
      start_date: startDate,
      end_date: endDate
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
        throw this.errorService.create(
          'Failed to fetch transactions from Plaid',
          { responseCode: responseCode, response: responseText, severity: 'high' }
        );
      }
      
      return JSON.parse(responseText);
    } catch (error) {
      this.errorService.handle(error, 'Failed to fetch transactions from Plaid');
      throw error;
    }
  };

  /**
   * Maps a Plaid category to the application's transaction type.
   * @private
   * @param {Array<string>} plaidCategories - Plaid category array.
   * @returns {string} The mapped transaction type.
   */
  PlaidServiceModuleConstructor.prototype._mapCategory = function(plaidCategories) {
    if (!plaidCategories || plaidCategories.length === 0) {
      return 'Extra';
    }
    
    const categoryMap = this.config.getSection('PLAID').CATEGORY_MAP || {};
    const primaryCategory = plaidCategories[0];
    
    return categoryMap[primaryCategory] || 'Extra';
  };

  /**
   * Imports Plaid transactions to the Transactions sheet.
   * @param {Array<object>} transactions - Array of Plaid transaction objects.
   * @returns {number} Number of transactions imported.
   * @memberof PlaidServiceModule
   */
  PlaidServiceModuleConstructor.prototype.importToSheet = function(transactions) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = this.config.getSheetNames();
    const transactionSheet = ss.getSheetByName(sheetNames.TRANSACTIONS);
    
    if (!transactionSheet) {
      throw this.errorService.create(
        'Transactions sheet not found',
        { severity: 'high' }
      );
    }
    
    // Transform Plaid transactions to sheet format: Date, Type, Category, Sub-Category, Shared?, Amount
    const dataToImport = transactions.map(function(tx) {
      return [
        new Date(tx.date),
        this._mapCategory(tx.category),
        tx.merchant_name || tx.name || 'Unknown',
        '', // Sub-category (to be filled by user)
        false, // Shared
        Math.abs(tx.amount) // Plaid amounts are negative for debits
      ];
    }.bind(this));
    
    if (dataToImport.length === 0) {
      return 0;
    }
    
    // Append to sheet
    const lastRow = transactionSheet.getLastRow();
    transactionSheet.getRange(lastRow + 1, 1, dataToImport.length, 6).setValues(dataToImport);
    
    return dataToImport.length;
  };

  return PlaidServiceModuleConstructor;
})();
