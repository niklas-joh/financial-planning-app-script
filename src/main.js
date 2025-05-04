/**
 * Financial Planning Tools - Main Entry Point
 * 
 * This file serves as the main entry point for the Financial Planning Tools
 * Google Apps Script project. It contains the onOpen function that sets up
 * all menu items for the various features.
 */

/**
 * Creates custom menus when the spreadsheet is opened
 * This combined function adds menu items for all features:
 * - Dropdown Tools
 * - Financial Overview
 * - Monthly Spending Report
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create single Financial Tools menu with all items
  ui.createMenu('Financial Tools')
    .addItem('Refresh Dropdown Cache', 'refreshCache')
    .addSeparator()
    .addItem('Generate Overview Sheet', 'createFinancialOverview')
    .addSeparator()
    .addItem('Generate Monthly Spending Report', 'generateMonthlySpendingReport')
    .addToUi();
}
