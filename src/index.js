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
  
  // Register global functions for Google Apps Script
  registerGlobalFunctions();
  
  // Log successful initialization
  Logger.log('Financial Planning Tools initialized successfully');
  Logger.log(`Version: ${FinancialPlanner.VERSION}`);
}

/**
 * Registers global functions that need to be accessible from Google Sheets UI
 * This is necessary because Google Apps Script requires global functions for triggers and menu items
 */
function registerGlobalFunctions() {
  // Core functions for Google Apps Script triggers
  const coreFunctions = [
    'onOpen',
    'onEdit'
  ];
  
  // Financial overview functions
  const overviewFunctions = [
    'createFinancialOverview',
    'handleOverviewSheetEdits'
  ];
  
  // Report functions
  const reportFunctions = [
    'generateMonthlySpendingReport',
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
    'toggleShowSubCategories',
    'setBudgetTargets',
    'setupEmailReports'
  ];
  
  // Utility functions
  const utilityFunctions = [
    'refreshCache'
  ];
  
  // Combine all function lists
  const allFunctions = [
    ...coreFunctions,
    ...overviewFunctions,
    ...reportFunctions,
    ...visualizationFunctions,
    ...financialAnalysisFunctions,
    ...settingsFunctions,
    ...utilityFunctions
  ];
  
  // Create global references to functions in the FinancialPlanner namespace
  allFunctions.forEach(funcName => {
    // Skip if function is already defined globally
    if (typeof this[funcName] !== 'undefined') {
      return;
    }
    
    // Find the function in the FinancialPlanner namespace
    // This will be updated as modules are refactored to use the namespace pattern
    if (FinancialPlanner.Controllers && typeof FinancialPlanner.Controllers[funcName] === 'function') {
      this[funcName] = FinancialPlanner.Controllers[funcName];
    }
    // Add more namespace checks as modules are refactored
  });
  
  // Log registration results
  Logger.log(`Registered ${allFunctions.length} global functions`);
}

// Run initialization
initialize();
