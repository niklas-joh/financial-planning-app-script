/**
 * @fileoverview Main entry point for the Financial Planning Tools Google Apps Script project.
 * This script handles the final initialization steps after all modules are loaded.
 * It ensures that the application's core components are ready and logs the
 * application version. It's designed to be one of the last scripts executed
 * as per the `appsscript.json` file order.
 * @module core/index
 */

/**
 * Initializes the Financial Planning Tools application.
 * This function logs the start of the initialization process, verifies that the
 * global `FinancialPlanner` namespace and its version are available (indicating
 * that `00_module_loader.js` has run), and then logs a success message along
 * with the application version.
 * It is called automatically when this script is loaded, provided the
 * `FinancialPlanner.Controllers` module is detected.
 * @memberof module:core/index
 * @private
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
