/**
 * @fileoverview SaltEdge Client Module - Consolidated SaltEdge AIS integration
 * Handles authentication, API communication, data fetching, and sheet operations
 * for SaltEdge Account Information Service integration.
 * @module services/salt-edge/salt-edge-client
 */

// Ensure the global FinancialPlanner namespace exists
var FinancialPlanner = FinancialPlanner || {};

/**
 * SaltEdge Client - Consolidated module for SaltEdge AIS integration
 * Follows KISS, DRY, YAGNI, and SRP principles with single module approach
 * @namespace FinancialPlanner.SaltEdgeClient
 */
FinancialPlanner.SaltEdgeClient = (function() {
  
  // Constants
  const API_URL = 'https://www.saltedge.com/api/v6';
  const CONSENT_SCOPES = ['accounts', 'transactions'];
  const FETCH_SCOPES = ['accounts', 'balance', 'transactions'];
  const DEFAULT_CUSTOMER_IDENTIFIER = 'financial_planner_user';
  
  /**
   * Gets SaltEdge API credentials from Script Properties
   * @returns {{appId: string, secret: string}} API credentials
   * @throws {Error} If credentials are not configured
   * @private
   */
  function getCredentials() {
    const props = PropertiesService.getScriptProperties();
    const appId = props.getProperty('SALTEDGE_APP_ID');
    const secret = props.getProperty('SALTEDGE_SECRET');
    
    if (!appId || !secret) {
      throw FinancialPlanner.ErrorService.create(
        'SaltEdge credentials not configured. Run Setup SaltEdge first.',
        { severity: 'high' }
      );
    }
    
    return { appId: appId, secret: secret };
  }

  /**
   * Gets cached private key from Script Properties
   * @returns {string} RSA private key in PEM format
   * @throws {Error} If private key is not cached
   * @private
   */
  function getPrivateKey() {
    const props = PropertiesService.getScriptProperties();
    const privateKey = props.getProperty('SALTEDGE_PRIVATE_KEY');
    
    if (!privateKey) {
      throw FinancialPlanner.ErrorService.create(
        'Private key not cached. Run Setup SaltEdge first.',
        { severity: 'high' }
      );
    }
    
    return privateKey;
  }

  /**
   * Generates RSA-SHA256 signature for SaltEdge API requests
   * Format: base64(sha256_signature(privateKey, "expires_at|method|url|body"))
   * @param {number} expiresAt - Unix timestamp (current time + 60 seconds)
   * @param {string} method - HTTP method (GET, POST, PUT, DELETE)
   * @param {string} url - Full request URL
   * @param {string} body - Request body (empty string for GET requests)
   * @returns {string} Base64 encoded RSA-SHA256 signature
   * @private
   */
  function generateSignature(expiresAt, method, url, body) {
    try {
      const privateKey = getPrivateKey();
      
      // Create signature string: expires_at|method|url|body
      const signatureString = expiresAt + '|' + method.toUpperCase() + '|' + url + '|' + (body || '');
      
      // Generate SHA256 signature using Apps Script built-in function
      const signature = Utilities.computeRsaSha256Signature(signatureString, privateKey);
      
      // Return base64 encoded signature
      return Utilities.base64Encode(signature);
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'Failed to generate SaltEdge API signature');
      throw error;
    }
  }

  /**
   * Makes authenticated request to SaltEdge API with proper signature
   * @param {string} endpoint - API endpoint (e.g., '/customers', '/connections')
   * @param {string} method - HTTP method (GET, POST, PUT, DELETE)
   * @param {Object} [params=null] - Request parameters (body for POST/PUT, query for GET)
   * @returns {Object} Parsed JSON response from SaltEdge API
   * @private
   */
  function makeRequest(endpoint, method, params) {
    try {
      const credentials = getCredentials();
      const expiresAt = Math.floor(Date.now() / 1000) + 60; // Current time + 60 seconds
      
      // Prepare request URL and body
      let fullUrl = API_URL + endpoint;
      let body = '';
      
      if (method === 'GET' && params) {
        // Append query parameters for GET requests
        const queryString = Object.keys(params)
          .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
          .join('&');
        fullUrl = fullUrl + '?' + queryString;
      } else if (params) {
        // JSON body for POST/PUT requests
        body = JSON.stringify(params);
      }
      
      // Generate signature
      const signature = generateSignature(expiresAt, method, fullUrl, body);
      
      // Prepare request headers
      const headers = {
        'App-id': credentials.appId,
        'Secret': credentials.secret,
        'Expires-at': expiresAt.toString(),
        'Signature': signature,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      };
      
      // Make HTTP request
      const options = {
        method: method.toLowerCase(),
        headers: headers,
        muteHttpExceptions: true
      };
      
      if (body) {
        options.payload = body;
      }
      
      const response = UrlFetchApp.fetch(fullUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      // Handle non-success responses
      if (responseCode !== 200 && responseCode !== 201) {
        let errorMessage = 'SaltEdge API request failed';
        try {
          const errorData = JSON.parse(responseText);
          if (errorData.error && errorData.error.message) {
            errorMessage = errorData.error.message;
          }
        } catch (parseError) {
          // Use default error message if response is not valid JSON
        }
        
        throw FinancialPlanner.ErrorService.create(errorMessage, {
          responseCode: responseCode,
          response: responseText,
          endpoint: endpoint,
          severity: 'high'
        });
      }
      
      return JSON.parse(responseText);
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'SaltEdge API request failed: ' + endpoint);
      throw error;
    }
  }

  /**
   * Flattens nested object into single-level with dot-notation keys
   * Reuses existing pattern from Plaid implementation for consistency
   * @param {Object} obj - Object to flatten
   * @param {string} [prefix=''] - Prefix for nested keys
   * @returns {Object} Flattened object with dot-notation keys
   * @private
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
        flat[newKey] = value.join(', ');
      } else if (typeof value === 'object' && !(value instanceof Date)) {
        Object.assign(flat, flattenObject(value, newKey));
      } else {
        flat[newKey] = value;
      }
    }
    
    return flat;
  }

  /**
   * Gets default from_date for data fetching (90 days ago)
   * @returns {string} Date in YYYY-MM-DD format
   * @private
   */
  function getDefaultFromDate() {
    const date = new Date();
    date.setDate(date.getDate() - 90);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  /**
   * Gets or creates a SaltEdge customer automatically
   * Stores customer ID in settings for future use
   * @param {string} [identifier] - Custom identifier, defaults to standard identifier
   * @returns {Object} Customer object with customer_id
   * @memberof FinancialPlanner.SaltEdgeClient
   */
  function getOrCreateCustomer(identifier) {
    identifier = identifier || DEFAULT_CUSTOMER_IDENTIFIER;
    
    try {
      // Check if customer already exists in settings
      let customerId = FinancialPlanner.SettingsService.getSaltEdgeCustomerId();
      
      if (customerId) {
        Logger.log('Using existing SaltEdge customer: ' + customerId);
        return { customer_id: customerId };
      }
      
      // Create new customer
      Logger.log('Creating new SaltEdge customer: ' + identifier);
      
      const response = makeRequest('/customers', 'POST', {
        data: {
          identifier: identifier
        }
      });
      
      customerId = response.data.customer_id;
      FinancialPlanner.SettingsService.setSaltEdgeCustomerId(customerId);
      
      Logger.log('Created SaltEdge customer: ' + customerId);
      return response.data;
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'Failed to get or create SaltEdge customer');
      throw error;
    }
  }

  /**
   * Creates connection URL for SaltEdge Widget
   * Automatically creates customer if needed
   * @param {Object} [options={}] - Connection options
   * @returns {{connect_url: string, expires_at: string, customer_id: string}} Widget connection data
   * @memberof FinancialPlanner.SaltEdgeClient
   */
  function createConnectionUrl(options) {
    options = options || {};
    
    try {
      // Get or create customer
      const customer = getOrCreateCustomer();
      const customerId = customer.customer_id;
      
      Logger.log('Creating SaltEdge connection URL for customer: ' + customerId);
      
      const payload = {
        data: {
          customer_id: customerId,
          consent: {
            scopes: options.consentScopes || CONSENT_SCOPES,
            from_date: options.fromDate || getDefaultFromDate(),
            period_days: options.periodDays || 90
          },
          attempt: {
            fetch_scopes: options.fetchScopes || FETCH_SCOPES,
            return_to: options.returnTo || ScriptApp.getService().getUrl()
          },
          widget: {
            show_consent_confirmation: true,
            skip_provider_selection: false,
            theme: 'default'
          }
        }
      };
      
      const response = makeRequest('/connections/connect', 'POST', payload);
      
      Logger.log('SaltEdge Widget URL created. Expires at: ' + response.data.expires_at);
      
      return {
        connect_url: response.data.connect_url,
        expires_at: response.data.expires_at,
        customer_id: response.data.customer_id
      };
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'Failed to create SaltEdge connection URL');
      throw error;
    }
  }

  /**
   * Lists all connections for the current customer
   * @returns {Array<Object>} Array of connection objects
   * @memberof FinancialPlanner.SaltEdgeClient
   */
  function listConnections() {
    try {
      const customerId = FinancialPlanner.SettingsService.getSaltEdgeCustomerId();
      
      if (!customerId) {
        throw FinancialPlanner.ErrorService.create(
          'No SaltEdge customer found. Connect a bank account first.',
          { severity: 'medium' }
        );
      }
      
      Logger.log('Fetching SaltEdge connections for customer: ' + customerId);
      
      const response = makeRequest('/connections', 'GET', {
        customer_id: customerId
      });
      
      Logger.log('Found ' + response.data.length + ' SaltEdge connections');
      return response.data;
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'Failed to list SaltEdge connections');
      throw error;
    }
  }

  /**
   * Lists all accounts for a specific connection
   * @param {string} connectionId - SaltEdge connection ID
   * @param {string} customerId - SaltEdge customer ID
   * @returns {Array<Object>} Array of account objects
   * @memberof FinancialPlanner.SaltEdgeClient
   */
  function listAccounts(connectionId, customerId) {
    try {
      Logger.log('Fetching SaltEdge accounts for connection: ' + connectionId);
      
      const response = makeRequest('/accounts', 'GET', {
        connection_id: connectionId,
        customer_id: customerId
      });
      
      Logger.log('Found ' + response.data.length + ' accounts');
      return response.data;
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'Failed to list SaltEdge accounts');
      throw error;
    }
  }

  /**
   * Lists all transactions for a specific account
   * @param {string} connectionId - SaltEdge connection ID
   * @param {string} accountId - SaltEdge account ID
   * @param {Object} [options={}] - Query options (pending, from_date, etc.)
   * @returns {Array<Object>} Array of transaction objects
   * @memberof FinancialPlanner.SaltEdgeClient
   */
  function listTransactions(connectionId, accountId, options) {
    options = options || {};
    
    try {
      Logger.log('Fetching SaltEdge transactions for account: ' + accountId);
      
      const params = {
        connection_id: connectionId,
        account_id: accountId
      };
      
      // Add optional filters
      if (options.pending !== undefined) {
        params.pending = options.pending;
      }
      
      const response = makeRequest('/transactions', 'GET', params);
      
      Logger.log('Found ' + response.data.length + ' transactions');
      return response.data;
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'Failed to list SaltEdge transactions');
      throw error;
    }
  }

  /**
   * Imports transactions to Google Sheet with dynamic column headers
   * Creates sheet if it doesn't exist, handles flattened transaction structure
   * @param {Array<Object>} transactions - Array of transaction objects
   * @returns {number} Number of transactions imported
   * @private
   */
  function importTransactionsToSheet(transactions) {
    if (!transactions || transactions.length === 0) {
      Logger.log('No transactions to import');
      return 0;
    }
    
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetName = FinancialPlanner.Config.getSheetNames().SALTEDGE_TRANSACTIONS;
      const sheet = FinancialPlanner.Utils.getOrCreateSheet(ss, sheetName);
      
      // Get or create headers from first transaction
      let headers;
      if (sheet.getLastRow() === 0) {
        // Create headers dynamically from first transaction
        const firstTx = transactions[0];
        const flattened = flattenObject(firstTx);
        headers = Object.keys(flattened);
        
        Logger.log('Creating SaltEdge sheet with ' + headers.length + ' dynamic columns');
        
        sheet.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight('bold')
          .setBackground('#4285F4')
          .setFontColor('#FFFFFF');
      } else {
        // Use existing headers
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        Logger.log('Using existing headers: ' + headers.length + ' columns');
      }
      
      // Prepare transaction rows
      const rows = transactions.map(function(tx) {
        const flattened = flattenObject(tx);
        return headers.map(function(header) {
          const value = flattened[header];
          // Parse date fields
          if ((header.includes('date') || header.includes('_at')) && value && typeof value === 'string') {
            try {
              return new Date(value);
            } catch (dateError) {
              return value;
            }
          }
          return value !== undefined ? value : '';
        });
      });
      
      // Append rows to sheet
      if (rows.length > 0) {
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
        
        // Format currency columns (amount fields)
        const amountColumns = headers.map((header, index) => 
          header.toLowerCase().includes('amount') || header.toLowerCase().includes('balance') ? index + 1 : null
        ).filter(col => col !== null);
        
        if (amountColumns.length > 0) {
          const currencyFormat = FinancialPlanner.Config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT;
          amountColumns.forEach(col => {
            sheet.getRange(lastRow + 1, col, rows.length, 1).setNumberFormat(currencyFormat);
          });
        }
      }
      
      Logger.log('Imported ' + transactions.length + ' transactions to SaltEdge sheet');
      return transactions.length;
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'Failed to import SaltEdge transactions to sheet');
      throw error;
    }
  }

  // Public API
  return {
    /**
     * One-time setup: prompts for RSA private key and API credentials, stores in Script Properties
     * @returns {string} Success message
     * @memberof FinancialPlanner.SaltEdgeClient
     */
    setup: function() {
      try {
        Logger.log('Starting SaltEdge setup...');
        
        const props = PropertiesService.getScriptProperties();
        const ui = SpreadsheetApp.getUi();
        
        // Check if already configured
        const existingAppId = props.getProperty('SALTEDGE_APP_ID');
        const existingSecret = props.getProperty('SALTEDGE_SECRET');
        const existingPrivateKey = props.getProperty('SALTEDGE_PRIVATE_KEY');
        
        if (existingAppId && existingSecret && existingPrivateKey) {
          return 'SaltEdge setup complete. All credentials already configured.';
        }
        
        // Prompt for RSA Private Key (if not already stored)
        if (!existingPrivateKey) {
          const privateKeyResponse = ui.prompt(
            'SaltEdge Setup - RSA Private Key',
            'Paste your RSA private key (complete PEM format including -----BEGIN/END----- lines):',
            ui.ButtonSet.OK_CANCEL
          );
          
          if (privateKeyResponse.getSelectedButton() !== ui.Button.OK) {
            throw FinancialPlanner.ErrorService.create('Setup cancelled by user', { severity: 'medium' });
          }
          
          const privateKey = privateKeyResponse.getResponseText().trim();
          if (!privateKey.includes('-----BEGIN') || !privateKey.includes('-----END')) {
            throw FinancialPlanner.ErrorService.create(
              'Invalid RSA private key format. Must include -----BEGIN/END----- lines.',
              { severity: 'high' }
            );
          }
          
          props.setProperty('SALTEDGE_PRIVATE_KEY', privateKey);
          Logger.log('RSA private key stored successfully');
        }
        
        // Prompt for App ID (if not already stored)
        if (!existingAppId) {
          const appIdResponse = ui.prompt(
            'SaltEdge Setup - App ID',
            'Enter your SaltEdge App-ID:',
            ui.ButtonSet.OK_CANCEL
          );
          
          if (appIdResponse.getSelectedButton() !== ui.Button.OK) {
            throw FinancialPlanner.ErrorService.create('Setup cancelled by user', { severity: 'medium' });
          }
          
          props.setProperty('SALTEDGE_APP_ID', appIdResponse.getResponseText().trim());
        }
        
        // Prompt for Secret (if not already stored)
        if (!existingSecret) {
          const secretResponse = ui.prompt(
            'SaltEdge Setup - Secret',
            'Enter your SaltEdge Secret:',
            ui.ButtonSet.OK_CANCEL
          );
          
          if (secretResponse.getSelectedButton() !== ui.Button.OK) {
            throw FinancialPlanner.ErrorService.create('Setup cancelled by user', { severity: 'medium' });
          }
          
          props.setProperty('SALTEDGE_SECRET', secretResponse.getResponseText().trim());
        }
        
        Logger.log('SaltEdge setup completed successfully');
        return 'SaltEdge setup complete. All credentials stored securely in Script Properties.';
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to complete SaltEdge setup');
        throw error;
      }
    },

    /**
     * Connects bank account via SaltEdge Widget
     * Auto-creates customer and displays Widget URL in modal dialog
     * @returns {string} Success message with Widget URL
     * @memberof FinancialPlanner.SaltEdgeClient
     */
    connectBank: function() {
      try {
        // Generate Widget URL (automatically creates customer if needed)
        const result = createConnectionUrl();
        
        // Show Widget URL in modal dialog
        const htmlOutput = HtmlService.createHtmlOutput(
          '<div style="padding: 20px; font-family: Arial, sans-serif;">' +
          '<h2 style="color: #4285F4; margin-bottom: 20px;">🏦 Connect Your Bank with SaltEdge</h2>' +
          '<p style="margin-bottom: 20px;">Click the link below to securely connect your bank account:</p>' +
          '<p style="margin-bottom: 30px;">' +
          '<a href="' + result.connect_url + '" target="_blank" ' +
          'style="display: inline-block; background: #4285F4; color: white; padding: 12px 24px; ' +
          'text-decoration: none; border-radius: 4px; font-size: 16px;">🔗 Open SaltEdge Widget</a>' +
          '</p>' +
          '<p style="color: #666; font-size: 12px; margin-bottom: 20px;">' +
          'Link expires at: ' + result.expires_at + '</p>' +
          '<p style="color: #666; font-size: 12px;">' +
          'After completing authentication, return here and use "Import SaltEdge" to fetch your data.' +
          '</p>' +
          '<button onclick="google.script.host.close()" ' +
          'style="background: #ccc; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer;">Close</button>' +
          '</div>'
        ).setWidth(500).setHeight(300);
        
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SaltEdge Bank Connection');
        
        return 'SaltEdge Widget opened. Complete authentication in the new window.';
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to connect SaltEdge bank account');
        throw error;
      }
    },

    /**
     * Imports all data (accounts and transactions) from all connections
     * Combined operation that fetches everything in one action
     * @returns {string} Success message with import summary
     * @memberof FinancialPlanner.SaltEdgeClient
     */
    importAllData: function() {
      try {
        Logger.log('Starting SaltEdge data import...');
        
        // Get all connections
        const connections = listConnections();
        
        if (connections.length === 0) {
          return 'No SaltEdge connections found. Connect a bank account first.';
        }
        
        let totalTransactions = 0;
        
        // Process each connection
        connections.forEach(function(connection) {
          Logger.log('Processing connection: ' + connection.provider_name);
          
          // Get accounts for this connection
          const accounts = listAccounts(connection.id, connection.customer_id);
          
          // Get transactions for each account
          accounts.forEach(function(account) {
            Logger.log('Processing account: ' + account.name);
            
            const transactions = listTransactions(connection.id, account.id);
            
            if (transactions.length > 0) {
              totalTransactions += importTransactionsToSheet(transactions);
            }
          });
        });
        
        const summary = 'Successfully imported ' + totalTransactions + ' transactions from ' + 
                       connections.length + ' SaltEdge connection(s)';
        
        Logger.log(summary);
        return summary;
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to import SaltEdge data');
        throw error;
      }
    },

    // Utility functions for external access if needed
    flattenObject: flattenObject,
    listConnections: listConnections
  };
})();
