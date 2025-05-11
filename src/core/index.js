/**
 * Financial Planning Tools - Index
 * 
 * This file serves as the main entry point for the Financial Planning Tools
 * Google Apps Script project. It imports all the necessary files and exposes
 * the public functions that can be called from the Google Sheets UI.
 * 
 * The file structure follows the namespace pattern to prevent global namespace
 * pollution and improve code organization.
 */

// Initialize the application
function initialize() {
  // Log initialization start
  Logger.log('Initializing Financial Planning Tools...');
  
  // Verify namespace is available
  if (typeof FinancialPlanner === 'undefined') {
    Logger.log('Error: FinancialPlanner namespace not defined. Make sure namespace.js is included first.');
    return;
  }
  
  // Log successful initialization
  Logger.log('Financial Planning Tools initialized successfully');
  Logger.log(`Version: ${FinancialPlanner.VERSION}`);
}

// Run initialization
// Ensure this runs after all modules in 00_module_loader.js are initialized.
// The 'files' array in appsscript.json should order index.js towards the end.
if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Controllers) {
  initialize();
} else {
  // This case should ideally not happen if file order is correct.
  // It means FinancialPlanner or its core components weren't loaded before index.js.
  Logger.log('FinancialPlanner namespace or Controllers not ready at the time of index.js execution. Initialization skipped.');
  // Consider a fallback or a way to defer initialize() if this becomes an issue.
}
