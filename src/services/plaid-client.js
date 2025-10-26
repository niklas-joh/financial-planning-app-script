/**
 * @fileoverview Plaid Client Module - Shared Plaid infrastructure.
 * Handles authentication, API configuration, and common utilities.
 * @module services/plaid-client
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Plaid Client - Shared Plaid authentication and utilities.
 * @namespace FinancialPlanner.PlaidClient
 */
FinancialPlanner.PlaidClient = (function() {
  /**
   * Gets the appropriate Plaid API URL based on environment setting.
   * @returns {string} The Plaid API base URL.
   * @memberof FinancialPlanner.PlaidClient
   */
  function getApiUrl() {
    const env = FinancialPlanner.SettingsService.getPlaidEnvironment();
    const envUrls = FinancialPlanner.Config.getSection('PLAID').ENVIRONMENTS;
    return env === 'production' ? envUrls.PRODUCTION : envUrls.SANDBOX;
  }

  /**
   * Gets API credentials from Script Properties for the current environment.
   * @returns {{clientId: string, secret: string}} The Plaid credentials.
   * @throws {Error} If credentials are not configured.
   * @memberof FinancialPlanner.PlaidClient
   */
  function getCredentials() {
    const env = FinancialPlanner.SettingsService.getPlaidEnvironment();
    const prefix = 'PLAID_' + env.toUpperCase() + '_';
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty(prefix + 'CLIENT_ID');
    const secret = props.getProperty(prefix + 'SECRET');
    
    if (!clientId || !secret) {
      throw FinancialPlanner.ErrorService.create(
        'Plaid credentials not configured for ' + env + ' environment. Please set ' + prefix + 'CLIENT_ID and ' + prefix + 'SECRET in Script Properties.',
        { severity: 'high' }
      );
    }
    
    return { clientId: clientId, secret: secret };
  }

  /**
   * Flattens a nested object into a single-level object with dot-notation keys.
   * Converts arrays to comma-separated strings.
   * @param {object} obj - The object to flatten.
   * @param {string} [prefix=''] - The prefix for nested keys.
   * @returns {object} The flattened object.
   * @memberof FinancialPlanner.PlaidClient
   */
  function flattenObject(obj, prefix) {
    prefix = prefix || '';
    const flat = {};
    
    for (const key in obj) {
      if (!obj.hasOwnProperty(key)) continue;
      
      const value = obj[key];
      const newKey = prefix ? prefix + '.' + key : key;
      
      if (value === null || value === undefined) {
        flat[newKey] = '';
      } else if (Array.isArray(value)) {
        // Convert arrays to comma-separated strings
        flat[newKey] = value.join(', ');
      } else if (typeof value === 'object' && !(value instanceof Date)) {
        // Recursively flatten nested objects
        Object.assign(flat, flattenObject(value, newKey));
      } else {
        flat[newKey] = value;
      }
    }
    
    return flat;
  }

  /**
   * Helper function to safely convert value, handling null/undefined.
   * @param {*} value - The value to convert.
   * @param {*} defaultValue - The default value if null/undefined.
   * @returns {*} The safe value.
   * @memberof FinancialPlanner.PlaidClient
   */
  function safeValue(value, defaultValue) {
    return (value !== null && value !== undefined) ? value : defaultValue;
  }

  // Public API
  return {
    getApiUrl: getApiUrl,
    getCredentials: getCredentials,
    flattenObject: flattenObject,
    safeValue: safeValue,

    /**
     * Creates a Plaid Link token for initiating the bank connection flow.
     * @returns {{link_token: string, expiration: string}} The link token response.
     * @memberof FinancialPlanner.PlaidClient
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
     * @memberof FinancialPlanner.PlaidClient
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
        
        // Store access token in Script Properties with environment prefix
        const env = FinancialPlanner.SettingsService.getPlaidEnvironment();
        const tokenKey = 'PLAID_' + env.toUpperCase() + '_ACCESS_TOKEN';
        PropertiesService.getScriptProperties().setProperty(tokenKey, result.access_token);
        
        return result;
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to exchange public token');
        throw error;
      }
    }
  };
})();
