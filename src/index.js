/**
 * Financial Planning Tools - Index
 * 
 * This file serves as the main entry point for the Financial Planning Tools
 * Google Apps Script project. It imports all the necessary files and exposes
 * the public functions that can be called from the Google Sheets UI.
 */

// Import utility functions
function importUtils() {
  // Common utility functions
  const commonUtils = [
    'columnToLetter',
    'getMonthName',
    'getOrCreateSheet',
    'formatAsCurrency',
    'formatAsPercentage',
    'setAlternatingRowColors'
  ];
  
  // Make utility functions globally available
  commonUtils.forEach(funcName => {
    this[funcName] = this[funcName] || global[funcName];
  });
}

// Import feature functions
function importFeatures() {
  // Dropdown feature functions
  const dropdownFunctions = [
    'onEdit',
    'refreshCache'
  ];
  
  // Financial overview functions
  const overviewFunctions = [
    'createFinancialOverview'
  ];
  
  // Monthly spending report functions
  const reportFunctions = [
    'generateMonthlySpendingReport'
  ];
  
  // Make feature functions globally available
  [...dropdownFunctions, ...overviewFunctions, ...reportFunctions].forEach(funcName => {
    this[funcName] = this[funcName] || global[funcName];
  });
}

// Initialize the application
function initialize() {
  importUtils();
  importFeatures();
}

// Run initialization
initialize();
