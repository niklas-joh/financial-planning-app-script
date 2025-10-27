# SaltEdge Integration Plan

## Overview
Implementation of SaltEdge Account Information Service (AIS) integration for the Financial Planning App, following KISS, DRY, YAGNI, and SRP principles.

## Refined MVP Architecture (Post-Critique)

### Key Design Decisions
1. **Single Module Approach**: One consolidated `salt-edge-client.js` file instead of 5 separate modules
2. **Existing Pattern Integration**: Reuse existing controllers.js menu structure and patterns
3. **Performance Optimization**: Cache private key in Script Properties instead of reading from Drive each time
4. **Minimal Menu Items**: Only 3 essential menu items for MVP
5. **Combined Operations**: Single "Import Data" function instead of separate account/transaction imports

### Architecture Summary

```
File Structure (Simplified):
src/services/salt-edge/
├── salt-edge-client.js     # Single consolidated module
├── public.pem              # RSA public key (existing)
└── private.pem             # RSA private key (existing)
```

### Core Features (MVP Scope)

✅ **Included**:
- Customer management (automatic creation)
- Connection via SaltEdge Widget
- Account and transaction fetching
- Import to dedicated "SaltEdge Transactions" sheet
- Manual refresh capability
- Basic error handling with existing ErrorService

❌ **Excluded (YAGNI)**:
- Holder info / KYC data
- Payment initiation (PIS)
- Automatic background refresh
- Transaction categorization
- Complex pagination handling
- Consent renewal automation

## Implementation Plan

### Phase 1: Documentation and Setup
- [x] Create comprehensive implementation plan
- [ ] Document SaltEdge configuration requirements
- [ ] Update system architecture documentation

### Phase 2: Core Implementation
- [ ] Implement consolidated salt-edge-client.js
- [ ] Update config.js with minimal SaltEdge settings
- [ ] Extend settings-service.js with SaltEdge methods
- [ ] Integrate with existing controllers.js menu system

### Phase 3: Testing and Integration
- [ ] Test with SaltEdge sandbox providers
- [ ] Validate data import to sheets
- [ ] Test error handling scenarios
- [ ] Document usage instructions

## Technical Specifications

### Configuration Changes

#### config.js (Minimal additions)
```javascript
SHEETS: {
  // ... existing sheets
  SALTEDGE_TRANSACTIONS: 'SaltEdge Transactions'  // Add only this
},

SALTEDGE: {
  API_URL: 'https://www.saltedge.com/api/v6',
  CONSENT_SCOPES: ['accounts', 'transactions'],
  FETCH_SCOPES: ['accounts', 'balance', 'transactions']
}
```

#### settings-service.js (Add methods)
```javascript
getSaltEdgeCustomerId: function() { /* implementation */ },
setSaltEdgeCustomerId: function(customerId) { /* implementation */ }
```

### Menu Integration (controllers.js)

#### Add to coreLogic object:
```javascript
const coreLogic = {
  // ... existing methods
  
  saltedgeSetup: function() {
    // One-time setup: cache private key + store credentials
  },
  
  saltedgeConnect: function() {
    // Auto-create customer, generate Widget URL, show in modal
  },
  
  saltedgeImport: function() {
    // Fetch all connections, accounts, and transactions in one operation
  }
};
```

#### Add to existing menu:
```javascript
.addSubMenu(ui.createMenu('🏦 Bank Integration')
  .addItem('🔗 Connect Bank Account', 'connectBankAccount_Global')
  // ... existing Plaid items
  .addSeparator()
  .addItem('⚙️ Setup SaltEdge', 'saltedgeSetup_Global')
  .addItem('🔗 Connect SaltEdge', 'saltedgeConnect_Global')
  .addItem('📥 Import SaltEdge', 'saltedgeImport_Global'))
```

### SaltEdge Client API Design

#### Core Functions:
```javascript
FinancialPlanner.SaltEdgeClient = {
  // Setup & Configuration
  setup: function() { /* One-time setup */ },
  
  // Authentication & Requests  
  makeRequest: function(endpoint, method, params) { /* API calls */ },
  
  // Customer Management
  getOrCreateCustomer: function(identifier) { /* Auto customer handling */ },
  
  // Connection Management
  createConnectionUrl: function(customerId) { /* Widget URL */ },
  listConnections: function(customerId) { /* All connections */ },
  
  // Data Operations
  listAccounts: function(connectionId, customerId) { /* Account data */ },
  listTransactions: function(connectionId, accountId) { /* Transaction data */ },
  
  // Sheet Operations
  importAllData: function() { /* Combined import */ },
  
  // Utilities
  flattenObject: function(obj) { /* Reuse existing pattern */ }
};
```

### Performance Optimizations

#### Private Key Caching:
```javascript
// Setup once (in saltedgeSetup):
function setup() {
  const privateKey = DriveApp.getFilesByName('private.pem')
    .next().getBlob().getDataAsString();
  PropertiesService.getScriptProperties()
    .setProperty('SALTEDGE_PRIVATE_KEY', privateKey);
}

// Use in signature generation:
function generateSignature() {
  const privateKey = PropertiesService.getScriptProperties()
    .getProperty('SALTEDGE_PRIVATE_KEY');
  // ... signature logic
}
```

### Error Handling Strategy

Leverage existing `FinancialPlanner.ErrorService`:
```javascript
try {
  // SaltEdge operation
} catch (error) {
  FinancialPlanner.ErrorService.handle(error, 'User-friendly message');
  throw error; // Re-throw for controller wrapper
}
```

Use existing `wrapWithFeedback` pattern in controllers.js.

### Sheet Integration

#### Dynamic Column Creation:
```javascript
function importTransactionsToSheet(transactions) {
  // Get or create "SaltEdge Transactions" sheet
  const sheet = FinancialPlanner.Utils.getOrCreateSheet(ss, sheetName);
  
  // Dynamic headers from first transaction
  const flattened = FinancialPlanner.SaltEdgeClient.flattenObject(transactions[0]);
  const headers = Object.keys(flattened);
  
  // Import with proper formatting
  // ... implementation
}
```

## User Experience Flow

### Setup (One-time):
1. User clicks "Setup SaltEdge"
2. System reads PEM files, prompts for credentials
3. Stores everything in Script Properties

### Connect Bank:
1. User clicks "Connect SaltEdge" 
2. System auto-creates customer if needed
3. Generates Widget URL, shows in modal dialog
4. User completes authentication in Widget

### Import Data:
1. User clicks "Import SaltEdge"
2. System fetches all connections
3. For each connection, fetches accounts and transactions
4. Imports to "SaltEdge Transactions" sheet with dynamic columns

### Error Scenarios:
- Missing credentials → Clear error message with setup instructions
- Widget authentication failure → User-friendly error dialog
- API rate limits → Handled gracefully with retry logic
- Invalid response data → Logged to Error Log sheet

## Security Considerations

1. **Private Key Storage**: Cached in Script Properties (encrypted by Google)
2. **API Credentials**: Stored in Script Properties (not in code)
3. **Request Signing**: All API requests signed with RSA-SHA256
4. **Environment Isolation**: Sandbox environment for testing

## Success Criteria

### Functionality:
- [ ] User can complete one-time setup
- [ ] User can connect SaltEdge bank account via Widget
- [ ] User can import accounts and transactions to sheet
- [ ] All operations handle errors gracefully

### Code Quality:
- [ ] Follows existing project patterns
- [ ] Single responsibility principle maintained
- [ ] DRY - reuses existing utilities and patterns
- [ ] KISS - simple, straightforward implementation
- [ ] YAGNI - only MVP features implemented

### Performance:
- [ ] Private key read only once (cached)
- [ ] API calls optimized (single connection fetch)
- [ ] Sheet operations use batch writing where possible

## Future Enhancements (Post-MVP)

1. **Automatic Refresh**: Background data sync
2. **Transaction Categorization**: Using SaltEdge's categorization API
3. **Holder Information**: KYC data integration
4. **Advanced Error Handling**: Retry logic, detailed error reporting
5. **Multi-Provider Support**: Handle multiple bank connections
6. **Data Validation**: Enhanced data quality checks

## Implementation Notes

### Dependencies:
- Existing: FinancialPlanner.Config, FinancialPlanner.ErrorService, FinancialPlanner.Utils
- New: Minimal additions to settings-service.js
- External: SaltEdge API v6, RSA signature generation

### Testing Strategy:
1. **Setup Testing**: Verify credential storage and private key caching
2. **Connection Testing**: Test Widget URL generation and customer creation  
3. **Data Import Testing**: Verify account and transaction import with fake providers
4. **Error Testing**: Test various failure scenarios

### Deployment Considerations:
- Ensure PEM files are properly deployed with clasp
- Test with SaltEdge sandbox environment first
- Verify all Script Properties are properly set
- Test menu integration and UI dialogs

This plan provides a complete roadmap for implementing a robust, maintainable SaltEdge integration that follows the project's architectural principles while delivering essential MVP functionality.
