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
    
    // Fix private key format: replace spaces with newlines for proper PEM structure
    return privateKey.replace(/\s+/g, '\n')
                     .replace(/-----BEGIN\n+PRIVATE\n+KEY-----/, '-----BEGIN PRIVATE KEY-----')
                     .replace(/-----END\n+PRIVATE\n+KEY-----/, '-----END PRIVATE KEY-----');
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
   * Gets list of stored connection IDs from Script Properties
   * @returns {Array<string>} Array of connection IDs
   * @private
   */
  function getStoredConnections() {
    const props = PropertiesService.getScriptProperties();
    const connectionsJson = props.getProperty('SALTEDGE_CONNECTIONS');
    return connectionsJson ? JSON.parse(connectionsJson) : [];
  }

  /**
   * Stores connection metadata in Script Properties and updates connections list
   * @param {Object} connectionData - Connection metadata to store
   * @param {string} connectionData.connection_id - Connection ID
   * @param {string} connectionData.customer_id - Customer ID
   * @param {string} connectionData.provider_name - Provider name
   * @param {string} connectionData.provider_code - Provider code
   * @param {string} connectionData.status - Connection status
   * @param {string} connectionData.created_at - Creation timestamp
   * @param {string} connectionData.last_synced_at - Last sync timestamp
   * @private
   */
  function storeConnection(connectionData) {
    const props = PropertiesService.getScriptProperties();
    const key = 'SALTEDGE_CONNECTION_' + connectionData.connection_id;
    props.setProperty(key, JSON.stringify(connectionData));
    
    // Update connections list
    const connections = getStoredConnections();
    if (!connections.includes(connectionData.connection_id)) {
      connections.push(connectionData.connection_id);
      props.setProperty('SALTEDGE_CONNECTIONS', JSON.stringify(connections));
    }
  }

  /**
   * Stores account metadata in Script Properties
   * @param {string} connectionId - Connection ID the account belongs to
   * @param {Object} accountData - Account metadata to store
   * @param {string} accountData.account_id - Account ID
   * @param {string} accountData.name - Account name
   * @param {string} accountData.nature - Account nature/type
   * @param {string} accountData.currency_code - Account currency code
   * @private
   */
  function storeAccount(connectionId, accountData) {
    const props = PropertiesService.getScriptProperties();
    const key = 'SALTEDGE_ACCOUNT_' + connectionId + '_' + accountData.account_id;
    const dataToStore = Object.assign({}, accountData, { connection_id: connectionId });
    props.setProperty(key, JSON.stringify(dataToStore));
  }

  /**
   * Gets stored pagination cursor for a specific account
   * @param {string} connectionId - Connection ID
   * @param {string} accountId - Account ID
   * @returns {string|null} Pagination cursor (next_id) or null if not set
   * @private
   */
  function getStoredCursor(connectionId, accountId) {
    const props = PropertiesService.getScriptProperties();
    const key = 'SALTEDGE_CURSOR_' + connectionId + '_' + accountId;
    return props.getProperty(key);
  }

  /**
   * Saves pagination cursor for a specific account
   * @param {string} connectionId - Connection ID
   * @param {string} accountId - Account ID
   * @param {string} nextId - Pagination cursor (next_id from API response)
   * @private
   */
  function saveCursor(connectionId, accountId, nextId) {
    const props = PropertiesService.getScriptProperties();
    const key = 'SALTEDGE_CURSOR_' + connectionId + '_' + accountId;
    if (nextId) {
      props.setProperty(key, nextId);
    }
  }

  /**
   * Removes connection and all associated data from Script Properties
   * Deletes connection metadata, accounts, and pagination cursors
   * @param {string} connectionId - Connection ID to remove
   * @private
   */
  function removeConnection(connectionId) {
    const props = PropertiesService.getScriptProperties();
    
    // Remove connection metadata
    props.deleteProperty('SALTEDGE_CONNECTION_' + connectionId);
    
    // Remove from connections list
    const connections = getStoredConnections();
    const updatedConnections = connections.filter(function(id) {
      return id !== connectionId;
    });
    props.setProperty('SALTEDGE_CONNECTIONS', JSON.stringify(updatedConnections));
    
    // Remove all accounts and cursors for this connection
    const allKeys = props.getKeys();
    allKeys.forEach(function(key) {
      if (key.startsWith('SALTEDGE_ACCOUNT_' + connectionId + '_') ||
          key.startsWith('SALTEDGE_CURSOR_' + connectionId + '_')) {
        props.deleteProperty(key);
      }
    });
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
   * Lists all transactions for a specific account with pagination support
   * Uses stored cursor for incremental imports (only fetches new transactions)
   * First import: Returns all transactions and saves cursor
   * Subsequent imports: Returns only new transactions since last cursor
   * @param {string} connectionId - SaltEdge connection ID
   * @param {string} accountId - SaltEdge account ID
   * @param {Object} [options={}] - Query options (pending, from_date, etc.)
   * @returns {Array<Object>} Array of transaction objects
   * @memberof FinancialPlanner.SaltEdgeClient
   */
  function listTransactions(connectionId, accountId, options) {
    options = options || {};
    
    try {
      let allTransactions = [];
      let nextId = getStoredCursor(connectionId, accountId);
      let hasMore = true;
      
      Logger.log('Fetching SaltEdge transactions for account: ' + accountId + 
                 (nextId ? ' (from cursor: ' + nextId + ')' : ' (full history)'));
      
      // Pagination loop to fetch all pages
      while (hasMore) {
        const params = {
          connection_id: connectionId,
          account_id: accountId,
          pending: false,      // Only posted transactions
          duplicated: false,   // Filter out duplicates
          per_page: 250        // Max per request
        };
        
        // Add pagination cursor if exists
        if (nextId) {
          params.from_id = nextId;
        }
        
        const response = makeRequest('/transactions', 'GET', params);
        
        allTransactions = allTransactions.concat(response.data);
        
        // Check for more pages
        nextId = response.meta ? response.meta.next_id : null;
        hasMore = !!nextId && response.data.length > 0;
        
        Logger.log('Fetched page: ' + response.data.length + ' transactions. Has more: ' + hasMore);
      }
      
      // Save final cursor for next import (null if no more data)
      if (allTransactions.length > 0 || nextId) {
        saveCursor(connectionId, accountId, nextId);
        Logger.log('Saved pagination cursor for next import: ' + (nextId || 'null (no more data)'));
      }
      
      Logger.log('Total transactions fetched: ' + allTransactions.length);
      return allTransactions;
    } catch (error) {
      FinancialPlanner.ErrorService.handle(error, 'Failed to list SaltEdge transactions');
      throw error;
    }
  }

  /**
   * Imports transactions to Google Sheet with metadata columns and dynamic headers
   * Prepends connection and account metadata at start of each row
   * Creates sheet if it doesn't exist, handles flattened transaction structure
   * @param {Array<Object>} transactions - Array of transaction objects
   * @param {Object} connectionMeta - Connection metadata (id, provider_name, etc.)
   * @param {Object} accountMeta - Account metadata (name, nature, currency_code, etc.)
   * @returns {number} Number of transactions imported
   * @private
   */
  function importTransactionsToSheet(transactions, connectionMeta, accountMeta) {
    if (!transactions || transactions.length === 0) {
      Logger.log('No transactions to import');
      return 0;
    }
    
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetName = FinancialPlanner.Config.getSheetNames().SALTEDGE_TRANSACTIONS;
      const sheet = FinancialPlanner.Utils.getOrCreateSheet(ss, sheetName);
      
      // Define metadata columns to prepend
      const metadataColumns = ['connection_id', 'provider_name', 'account_name'];
      
      // Get or create headers
      let headers;
      if (sheet.getLastRow() === 0) {
        // Create headers: metadata + transaction fields
        const firstTx = transactions[0];
        const flattened = flattenObject(firstTx);
        const txHeaders = Object.keys(flattened);
        
        headers = metadataColumns.concat(txHeaders);
        
        Logger.log('Creating SaltEdge sheet with ' + headers.length + ' columns (3 metadata + ' + txHeaders.length + ' transaction fields)');
        
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
      
      // Prepare transaction rows with metadata
      const rows = transactions.map(function(tx) {
        const flattened = flattenObject(tx);
        
        return headers.map(function(header) {
          // Prepend metadata values
          if (header === 'connection_id') {
            return connectionMeta.id;
          }
          if (header === 'provider_name') {
            return connectionMeta.provider_name;
          }
          if (header === 'account_name') {
            return accountMeta.name;
          }
          
          // Map transaction fields
          const value = flattened[header];
          
          // Parse date fields (reuse Plaid pattern)
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
        const amountColumns = headers.map(function(header, index) {
          return header.toLowerCase().includes('amount') || header.toLowerCase().includes('balance') ? index + 1 : null;
        }).filter(function(col) {
          return col !== null;
        });
        
        if (amountColumns.length > 0) {
          const currencyFormat = FinancialPlanner.Config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT;
          amountColumns.forEach(function(col) {
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
          'After completing authentication in the SaltEdge widget, close that window and return here. ' +
          'Then use "Financial Tools → Bank Integration → Import SaltEdge Data" to fetch your data.' +
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
     * Disconnects a SaltEdge connection and removes all associated data
     * Shows confirmation dialog, calls API to disconnect, removes local storage
     * @param {string} connectionId - Connection ID to disconnect
     * @returns {string} Success message
     * @memberof FinancialPlanner.SaltEdgeClient
     */
    disconnectConnection: function(connectionId) {
      try {
        const props = PropertiesService.getScriptProperties();
        const ui = SpreadsheetApp.getUi();
        
        // Get connection metadata for confirmation
        const connKey = 'SALTEDGE_CONNECTION_' + connectionId;
        const connDataJson = props.getProperty(connKey);
        
        if (!connDataJson) {
          throw FinancialPlanner.ErrorService.create(
            'Connection not found: ' + connectionId,
            { severity: 'medium' }
          );
        }
        
        const connData = JSON.parse(connDataJson);
        
        // Show confirmation dialog
        const confirmResponse = ui.alert(
          'Disconnect SaltEdge Account',
          'Are you sure you want to disconnect "' + connData.provider_name + '"?\\n\\n' +
          'This will remove the connection and all associated account metadata. ' +
          'Transaction data in the sheet will NOT be deleted.',
          ui.ButtonSet.YES_NO
        );
        
        if (confirmResponse !== ui.Button.YES) {
          return 'Disconnect cancelled by user';
        }
        
        Logger.log('Disconnecting SaltEdge connection: ' + connectionId);
        
        // Call SaltEdge API to disconnect
        try {
          makeRequest('/connections/' + connectionId, 'DELETE', null);
          Logger.log('Connection disconnected from SaltEdge API');
        } catch (apiError) {
          // Continue with local cleanup even if API call fails
          Logger.log('Warning: API disconnect failed, proceeding with local cleanup: ' + apiError.message);
        }
        
        // Remove connection and all associated data from Script Properties
        removeConnection(connectionId);
        
        const message = 'Successfully disconnected "' + connData.provider_name + '". ' +
                       'Transaction data remains in sheet.';
        
        Logger.log(message);
        return message;
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to disconnect SaltEdge connection');
        throw error;
      }
    },

    /**
     * Shows connected SaltEdge accounts in a modal dialog
     * Reads from Script Properties and displays in HTML table
     * @returns {void}
     * @memberof FinancialPlanner.SaltEdgeClient
     */
    showConnectedAccounts: function() {
      try {
        const connectionIds = getStoredConnections();
        
        if (connectionIds.length === 0) {
          const html = HtmlService.createHtmlOutput(
            '<div style="padding: 20px; font-family: Arial, sans-serif;">' +
            '<p>No SaltEdge accounts connected yet.</p>' +
            '<p style="margin-top: 15px; color: #666; font-size: 14px;">Use "Financial Tools → Bank Integration → Connect SaltEdge" to add a bank account.</p>' +
            '</div>'
          ).setWidth(400).setHeight(150);
          
          SpreadsheetApp.getUi().showModalDialog(html, 'Connected SaltEdge Accounts');
          return;
        }
        
        const props = PropertiesService.getScriptProperties();
        let tableRows = '';
        
        connectionIds.forEach(function(connId) {
          const connData = JSON.parse(props.getProperty('SALTEDGE_CONNECTION_' + connId));
          
          // Get accounts for this connection
          const allKeys = props.getKeys();
          allKeys.forEach(function(key) {
            if (key.startsWith('SALTEDGE_ACCOUNT_' + connId + '_')) {
              const accountData = JSON.parse(props.getProperty(key));
              tableRows += '<tr>' +
                '<td style="padding: 8px; border-bottom: 1px solid #ddd;">' + connData.provider_name + '</td>' +
                '<td style="padding: 8px; border-bottom: 1px solid #ddd;">' + accountData.name + '</td>' +
                '<td style="padding: 8px; border-bottom: 1px solid #ddd;">' + accountData.nature + '</td>' +
                '<td style="padding: 8px; border-bottom: 1px solid #ddd;">' + accountData.currency_code + '</td>' +
                '<td style="padding: 8px; border-bottom: 1px solid #ddd;">' + connData.status + '</td>' +
                '</tr>';
            }
          });
        });
        
        const html = HtmlService.createHtmlOutput(
          '<div style="padding: 20px; font-family: Arial, sans-serif;">' +
          '<h2 style="color: #4285F4; margin-top: 0;">Connected SaltEdge Accounts</h2>' +
          '<table style="border-collapse: collapse; width: 100%; margin-top: 20px;">' +
          '<thead><tr style="background: #4285F4; color: white;">' +
          '<th style="padding: 10px; text-align: left;">Provider</th>' +
          '<th style="padding: 10px; text-align: left;">Account</th>' +
          '<th style="padding: 10px; text-align: left;">Type</th>' +
          '<th style="padding: 10px; text-align: left;">Currency</th>' +
          '<th style="padding: 10px; text-align: left;">Status</th>' +
          '</tr></thead>' +
          '<tbody>' + tableRows + '</tbody>' +
          '</table>' +
          '<p style="margin-top: 20px; color: #666; font-size: 12px;">Use "Import SaltEdge Data" to fetch the latest transactions from these accounts.</p>' +
          '</div>'
        ).setWidth(700).setHeight(400);
        
        SpreadsheetApp.getUi().showModalDialog(html, 'Connected SaltEdge Accounts');
      } catch (error) {
        FinancialPlanner.ErrorService.handle(error, 'Failed to show connected accounts');
        throw error;
      }
    },

    /**
     * Imports all data (accounts and transactions) from all connections
     * Stores connection and account metadata, uses pagination for transactions
     * Combined operation that fetches everything in one action
     * @returns {string} Success message with import summary
     * @memberof FinancialPlanner.SaltEdgeClient
     */
    importAllData: function() {
      try {
        Logger.log('Starting SaltEdge data import...');
        
        // Get all connections from API
        const connections = listConnections();
        
        if (connections.length === 0) {
          return 'No SaltEdge connections found. Connect a bank account first.';
        }
        
        Logger.log('Found ' + connections.length + ' connection(s)');
        
        let totalTransactions = 0;
        
        // Process each connection
        connections.forEach(function(connection) {
          Logger.log('Processing connection: ' + connection.provider_name);
          
          // Store connection metadata in Script Properties
          storeConnection({
            connection_id: connection.id,
            customer_id: connection.customer_id,
            provider_name: connection.provider_name,
            provider_code: connection.provider_code,
            status: connection.status,
            created_at: connection.created_at,
            last_synced_at: new Date().toISOString()
          });
          
          // Get accounts for this connection
          const accounts = listAccounts(connection.id, connection.customer_id);
          
          Logger.log('Found ' + accounts.length + ' account(s) for connection: ' + connection.id);
          
          // Process each account
          accounts.forEach(function(account) {
            Logger.log('Processing account: ' + account.name);
            
            // Store account metadata in Script Properties
            storeAccount(connection.id, {
              account_id: account.id,
              name: account.name,
              nature: account.nature,
              currency_code: account.currency_code
            });
            
            // Fetch transactions with pagination (uses stored cursor)
            const transactions = listTransactions(connection.id, account.id);
            
            if (transactions.length > 0) {
              // Import with metadata
              totalTransactions += importTransactionsToSheet(
                transactions,
                connection,
                account
              );
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
