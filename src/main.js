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
  
  // Create enhanced Financial Tools menu with submenus and icons
  ui.createMenu('📊 Financial Tools')
    .addItem('📈 Generate Overview', 'createFinancialOverview')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Reports')
      .addItem('📝 Monthly Spending Report', 'generateMonthlySpendingReport')
      .addItem('📅 Yearly Summary (Coming Soon)', 'generateYearlySummary')
      .addItem('🔍 Category Breakdown (Coming Soon)', 'generateCategoryBreakdown')
      .addItem('💰 Savings Analysis (Coming Soon)', 'generateSavingsAnalysis'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Visualizations (Coming Soon)')
      .addItem('📉 Spending Trends Chart (Coming Soon)', 'createSpendingTrendsChart')
      .addItem('⚖️ Budget vs Actual (Coming Soon)', 'createBudgetVsActualChart')
      .addItem('💹 Income vs Expenses (Coming Soon)', 'createIncomeVsExpensesChart')
      .addItem('🍩 Category Pie Chart (Coming Soon)', 'createCategoryPieChart'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🧮 Financial Analysis')
      .addItem('💡 Suggest Savings Opportunities (Coming Soon)', 'suggestSavingsOpportunities')
      .addItem('⚠️ Spending Anomaly Detection (Coming Soon)', 'detectSpendingAnomalies')
      .addItem('📌 Fixed vs Variable Expenses (Coming Soon)', 'analyzeFixedVsVariableExpenses')
      .addItem('🔮 Cash Flow Forecast (Coming Soon)', 'generateCashFlowForecast'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Settings')
      .addItem('🔄 Toggle Sub-Categories in Overview', 'toggleShowSubCategories')
      .addItem('🎯 Set Budget Targets (Coming Soon)', 'setBudgetTargets')
      .addItem('📧 Setup Email Reports (Coming Soon)', 'setupEmailReports'))
    .addToUi();
}

/**
 * Event handler that runs when a user edits the spreadsheet.
 * Used to detect changes to settings checkboxes and other interactive elements.
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  // Pass the edit event to various handlers
  handleOverviewSheetEdits(e);
  
  // More handlers can be added here in the future
}
