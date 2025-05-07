/**
 * Financial Planning Tools - Main Entry Point
 * 
 * This file serves as the main entry point for the Financial Planning Tools
 * Google Apps Script project. It contains global function references that
 * delegate to the appropriate modules in the FinancialPlanner namespace.
 * 
 * The actual implementation of these functions is in the Controllers module.
 */

/**
 * Creates custom menus when the spreadsheet is opened
 * This function is automatically called by Google Apps Script when the spreadsheet is opened.
 * It delegates to the Controllers module in the FinancialPlanner namespace.
 */
function onOpen() {
  // Delegate to the Controllers module
  return FinancialPlanner.Controllers.onOpen();
}

/**
 * Event handler that runs when a user edits the spreadsheet.
 * Used to detect changes to settings checkboxes and other interactive elements.
 * This function is automatically called by Google Apps Script when the spreadsheet is edited.
 * It delegates to the Controllers module in the FinancialPlanner namespace.
 * 
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  // Delegate to the Controllers module
  return FinancialPlanner.Controllers.onEdit(e);
}

/**
 * Creates a financial overview sheet based on transaction data
 * This function delegates to the appropriate module in the FinancialPlanner namespace.
 * It is exposed globally for backward compatibility and for use in the UI.
 */
function createFinancialOverview() {
  // Delegate to the Controllers module
  return FinancialPlanner.Controllers.createFinancialOverview();
}

/**
 * Generates a monthly spending report
 * This function delegates to the appropriate module in the FinancialPlanner namespace.
 * It is exposed globally for backward compatibility and for use in the UI.
 */
function generateMonthlySpendingReport() {
  // Delegate to the Controllers module
  return FinancialPlanner.Controllers.generateMonthlySpendingReport();
}

/**
 * Shows key financial metrics
 * This function delegates to the appropriate module in the FinancialPlanner namespace.
 * It is exposed globally for backward compatibility and for use in the UI.
 */
function showKeyMetrics() {
  // Delegate to the Controllers module
  return FinancialPlanner.Controllers.showKeyMetrics();
}

/**
 * Toggles the display of sub-categories in the overview
 * This function delegates to the appropriate module in the FinancialPlanner namespace.
 * It is exposed globally for backward compatibility and for use in the UI.
 */
function toggleShowSubCategories() {
  // Delegate to the Controllers module
  return FinancialPlanner.Controllers.toggleShowSubCategories();
}

/**
 * Refreshes the cache
 * This function delegates to the appropriate module in the FinancialPlanner namespace.
 * It is exposed globally for backward compatibility and for use in the UI.
 */
function refreshCache() {
  // Delegate to the Controllers module
  return FinancialPlanner.Controllers.refreshCache();
}
