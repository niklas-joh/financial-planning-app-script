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
    // Use this[funcName] to assign to the global scope in Apps Script
    // Avoid using global[funcName] which causes ReferenceError
    if (typeof this[funcName] === 'undefined') {
      Logger.log('Warning: ' + funcName + ' is not defined');
    }
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
    'generateMonthlySpendingReport',
    // New report functions
    'generateYearlySummary',
    'generateCategoryBreakdown',
    'generateSavingsAnalysis'
  ];
  
  // Visualization functions
  const visualizationFunctions = [
    'createSpendingTrendsChart',
    'createBudgetVsActualChart',
    'createIncomeVsExpensesChart',
    'createCategoryPieChart'
  ];
  
  // Financial analysis functions
  const financialAnalysisFunctions = [
    'showKeyMetrics',
    'suggestSavingsOpportunities',
    'detectSpendingAnomalies',
    'analyzeFixedVsVariableExpenses',
    'generateCashFlowForecast'
  ];
  
  // Settings functions
  const settingsFunctions = [
    'setBudgetTargets',
    'setupEmailReports'
  ];
  
  // Make feature functions globally available
  [
    ...dropdownFunctions, 
    ...overviewFunctions, 
    ...reportFunctions,
    ...visualizationFunctions,
    ...financialAnalysisFunctions,
    ...settingsFunctions
  ].forEach(funcName => {
    // Use this[funcName] to assign to the global scope in Apps Script
    // Avoid using global[funcName] which causes ReferenceError
    if (typeof this[funcName] === 'undefined') {
      Logger.log('Warning: ' + funcName + ' is not defined');
    }
  });
}

// Initialize the application
function initialize() {
  importUtils();
  importFeatures();
  Logger.log('Financial Planning Tools initialized successfully');
}

// Run initialization
initialize();
