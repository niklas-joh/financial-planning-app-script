# SaltEdge Multiple Bank Account Support - Implementation Plan

## Problem Statement

When connecting multiple bank accounts via SaltEdge, transactions are being overwritten instead of being combined. The current implementation lacks:
- Transaction identification/deduplication logic
- Connection and account metadata tracking
- Pagination cursor storage for incremental imports
- User visibility into connected accounts

## Root Cause Analysis

### Current Issues

1. **No Source Identification**: `importTransactionsToSheet()` doesn't add metadata to identify which connection/account transactions belong to
2. **No Deduplication**: Unlike Plaid (which uses `transaction_id`), SaltEdge has no mechanism to prevent duplicate imports
3. **Missing Connection Registry**: Only `SaltEdgeCustomerId` is stored - no tracking of connections or accounts
4. **Overwriting Data**: Subsequent imports append without checking for duplicates, causing data corruption

## Solution Architecture (REVISED)

### Key Design Principles
- **KISS**: Use existing Script Properties pattern (like Plaid)
- **YAGNI**: No separate account overview sheet - use UI dialog
- **SRP**: SaltEdge Client manages its own storage
- **DRY**: Reuse Plaid patterns for pagination and storage

### Storage Strategy: Script Properties

**Why Script Properties over Settings Sheet:**
- Settings sheet is for user preferences (ShowSubCategories, PlaidEnvironment)
- Script Properties already used for Plaid cursors/tokens
- Atomic updates, no JSON parsing overhead
- Follows existing codebase pattern
- Maintains SRP - SaltEdge Client manages its own data

**Storage Structure:**
```javascript
// Script Properties keys:
SALTEDGE_CUSTOMER_ID = "customer_abc"
SALTEDGE_CONNECTIONS = ["conn1", "conn2"]  // JSON array of connection IDs

// Individual connection metadata
SALTEDGE_CONNECTION_conn1 = {
  "connection_id": "conn1",
  "customer_id": "customer_abc", 
  "provider_name": "Deutsche Bank",
  "provider_code": "deutsche_bank_de",
  "status": "active",
  "created_at": "2025-01-09T12:00:00Z",
  "last_synced_at": "2025-01-09T12:30:00Z"
}

// Individual account metadata
SALTEDGE_ACCOUNT_conn1_acc1 = {
  "account_id": "acc1",
  "connection_id": "conn1",
  "name": "Main Checking",
  "nature": "account",
  "currency_code": "EUR"
}

// Pagination cursors per account
SALTEDGE_CURSOR_conn1_acc1 = "3333333333333333"  // next_id from API
```

### Transaction Sheet Structure

**Enhanced Columns (prepended at start):**
```
| connection_id | provider_name | account_name | id | account_id | amount | date | ... (existing API fields) |
```

**Key Points:**
- Add ONLY 3 metadata columns (connection_id, provider_name, account_name)
- `account_id` already exists in SaltEdge transaction data
- Place metadata columns at START for visibility
- Reuse `flattenObject()` for dynamic column creation
- Apply same date parsing pattern as Plaid

### Pagination-Based Import (No Deduplication Needed!)

**SaltEdge List Transactions API:**
```
GET /transactions?connection_id={id}&account_id={id}&from_id={next_id}&per_page=250
```

**Response includes pagination:**
```json
{
  "data": [...],
  "meta": {
    "next_id": "3333333333333333",
    "next_page": "/api/v6/transactions?from_id=..."
  }
}
```

**Flow:**
1. First import: No `from_id` → Returns ALL transactions + `next_id`
2. Store `next_id` as cursor for account
3. Subsequent imports: Use stored `from_id` → Returns only NEW transactions
4. Update cursor with new `next_id`

**Result: Automatic deduplication via API pagination!**

## Implementation Subtasks

### Subtask 1: Storage Helpers in SaltEdge Client
Add private functions for Script Properties management:

```javascript
// Private helper functions
function getStoredConnections() {
  const props = PropertiesService.getScriptProperties();
  const connectionsJson = props.getProperty('SALTEDGE_CONNECTIONS');
  return connectionsJson ? JSON.parse(connectionsJson) : [];
}

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

function storeAccount(connectionId, accountData) {
  const props = PropertiesService.getScriptProperties();
  const key = 'SALTEDGE_ACCOUNT_' + connectionId + '_' + accountData.account_id;
  const dataToStore = Object.assign({}, accountData, { connection_id: connectionId });
  props.setProperty(key, JSON.stringify(dataToStore));
}

function getStoredCursor(connectionId, accountId) {
  const props = PropertiesService.getScriptProperties();
  const key = 'SALTEDGE_CURSOR_' + connectionId + '_' + accountId;
  return props.getProperty(key);
}

function saveCursor(connectionId, accountId, nextId) {
  const props = PropertiesService.getScriptProperties();
  const key = 'SALTEDGE_CURSOR_' + connectionId + '_' + accountId;
  if (nextId) {
    props.setProperty(key, nextId);
  }
}

function removeConnection(connectionId) {
  const props = PropertiesService.getScriptProperties();
  
  // Remove connection metadata
  props.deleteProperty('SALTEDGE_CONNECTION_' + connectionId);
  
  // Remove from connections list
  const connections = getStoredConnections();
  const updatedConnections = connections.filter(id => id !== connectionId);
  props.setProperty('SALTEDGE_CONNECTIONS', JSON.stringify(updatedConnections));
  
  // Remove all accounts and cursors for this connection
  const allKeys = props.getKeys();
  allKeys.forEach(key => {
    if (key.startsWith('SALTEDGE_ACCOUNT_' + connectionId + '_') ||
        key.startsWith('SALTEDGE_CURSOR_' + connectionId + '_')) {
      props.deleteProperty(key);
    }
  });
}
```

### Subtask 2: Paginated Transaction Fetching
Modify `listTransactions()` to support pagination with cursor:

```javascript
function listTransactions(connectionId, accountId, options) {
  options = options || {};
  
  let allTransactions = [];
  let nextId = getStoredCursor(connectionId, accountId);
  let hasMore = true;
  
  Logger.log('Fetching transactions for account: ' + accountId + 
             (nextId ? ' (from cursor)' : ' (full history)'));
  
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
  if (allTransactions.length > 0) {
    saveCursor(connectionId, accountId, nextId);
  }
  
  Logger.log('Total transactions fetched: ' + allTransactions.length);
  return allTransactions;
}
```

### Subtask 3: Enhanced Transaction Import
Modify `importTransactionsToSheet()` to accept and prepend metadata:

```javascript
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
      
      Logger.log('Creating SaltEdge sheet with ' + headers.length + ' columns');
      
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
      
      // Format currency columns
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
```

### Subtask 4: Update importAllData() Flow
Modify to store metadata and use pagination:

```javascript
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
      
      // Store connection metadata
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
        
        // Store account metadata
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
}
```

### Subtask 5: UI Dialog for Account Overview
Add public method to show connected accounts:

```javascript
showConnectedAccounts: function() {
  try {
    const connectionIds = getStoredConnections();
    
    if (connectionIds.length === 0) {
      const html = HtmlService.createHtmlOutput(
        '<div style="padding: 20px; font-family: Arial;">No accounts connected yet.</div>'
      ).setWidth(400).setHeight(200);
      
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
            '<td>' + connData.provider_name + '</td>' +
            '<td>' + accountData.name + '</td>' +
            '<td>' + accountData.nature + '</td>' +
            '<td>' + accountData.currency_code + '</td>' +
            '<td>' + connData.status + '</td>' +
            '</tr>';
        }
      });
    });
    
    const html = HtmlService.createHtmlOutput(
      '<div style="padding: 20px; font-family: Arial;">' +
      '<h2>Connected SaltEdge Accounts</h2>' +
      '<table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%;">' +
      '<thead><tr style="background: #4285F4; color: white;">' +
      '<th>Provider</th><th>Account</th><th>Type</th><th>Currency</th><th>Status</th>' +
      '</tr></thead>' +
      '<tbody>' + tableRows + '</tbody>' +
      '</table>' +
      '</div>'
    ).setWidth(700).setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Connected SaltEdge Accounts');
  } catch (error) {
    FinancialPlanner.ErrorService.handle(error, 'Failed to show connected accounts');
    throw error;
  }
}
```

### Subtask 6: Disconnect Account Functionality
Add public method with user confirmation:

```javascript
disconnectAccount: function() {
  try {
    const ui = SpreadsheetApp.getUi();
    const connectionIds = getStoredConnections();
    
    if (connectionIds.length === 0) {
      ui.alert('No accounts connected');
      return;
    }
    
    // Build list of connections for user selection
    const props = PropertiesService.getScriptProperties();
    let message = 'Select connection to disconnect:\n\n';
    
    connectionIds.forEach(function(connId, index) {
      const connData = JSON.parse(props.getProperty('SALTEDGE_CONNECTION_' + connId));
      message += (index + 1) + '. ' + connData.provider_name + ' (ID: ' + connId + ')\n';
    });
    
    const response = ui.prompt('Disconnect Account', message + '\nEnter number:', ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return 'Cancelled';
    }
    
    const selection = parseInt(response.getResponseText()) - 1;
    
    if (isNaN(selection) || selection < 0 || selection >= connectionIds.length) {
      ui.alert('Invalid selection');
      return;
    }
    
    const connectionId = connectionIds[selection];
    const connData = JSON.parse(props.getProperty('SALTEDGE_CONNECTION_' + connectionId));
    
    // Ask user about transaction history
    const keepHistoryResponse = ui.alert(
      'Keep Transaction History?',
      'Do you want to keep the imported transactions from ' + connData.provider_name + '?\n\n' +
      'YES = Keep transactions in sheet\n' +
      'NO = This option is not available (transactions remain)',
      ui.ButtonSet.YES_NO
    );
    
    // Remove connection metadata
    removeConnection(connectionId);
    
    ui.alert('Successfully disconnected ' + connData.provider_name);
    
    return 'Disconnected: ' + connData.provider_name;
  } catch (error) {
    FinancialPlanner.ErrorService.handle(error, 'Failed to disconnect account');
    throw error;
  }
}
```

### Subtask 7: Controller Integration
Add menu items to `controllers.js`:

```javascript
// In coreLogic object, add:
saltedgeShowAccounts: function() {
  FinancialPlanner.SaltEdgeClient.showConnectedAccounts();
},

saltedgeDisconnect: function() {
  FinancialPlanner.SaltEdgeClient.disconnectAccount();
}

// In createCustomMenu, add to Bank Integration submenu:
.addItem('📋 Show SaltEdge Accounts', 'saltedgeShowAccounts_Global')
.addItem('🔌 Disconnect SaltEdge Account', 'saltedgeDisconnect_Global')
```

## Review Points & Validation

### Review Points to Watch For:
1. **Cursor Storage**: Ensure `next_id` is saved AFTER successful import, not before
2. **Connection Removal**: Verify all related properties (accounts, cursors) are deleted
3. **Metadata Columns**: Confirm they appear at START of sheet, not end
4. **Date Parsing**: Apply same pattern as Plaid for consistency
5. **Error Handling**: Wrap all operations in try/catch with ErrorService

### Testing Checklist:
1. ✅ First import with no cursor → Returns all transactions
2. ✅ Cursor saved after import
3. ✅ Second import uses cursor → Returns only new transactions
4. ✅ Multiple connections/accounts work independently
5. ✅ Show accounts dialog displays correct data
6. ✅ Disconnect removes all metadata cleanly
7. ✅ Transaction history preserved in sheet

## Dependent Files

### Files to Modify:
1. `src/services/salt-edge/salt-edge-client.js` - All subtasks 1-6
2. `src/core/controllers.js` - Subtask 7 (menu integration)

### Files Referenced (no changes):
1. `src/core/config.js` - Read SALTEDGE_TRANSACTIONS sheet name
2. `src/services/settings-service.js` - Use existing getSaltEdgeCustomerId/setSaltEdgeCustomerId
3. `src/services/error-service.js` - Error handling
4. `src/utils/common.js` - getOrCreateSheet utility

## Benefits Summary

✅ **DRY**: Reuses Plaid patterns for storage and pagination  
✅ **SRP**: SaltEdge Client manages its own Script Properties  
✅ **KISS**: No unnecessary sheets, direct property access  
✅ **YAGNI**: No over-engineered account overview sheet  
✅ **Performance**: Pagination eliminates duplicate checking overhead  
✅ **Consistency**: Matches existing codebase patterns  
✅ **Maintainability**: Clear, simple, follows development principles

## Next Steps

1. Implement Subtask 1: Storage helpers
2. Review code for consistency, logic, potential flaws
3. Analyze impact on other files
4. Request user validation
5. Update this documentation
6. Git commit with comprehensive message
7. Create new Cline task for next subtask (repeat)
