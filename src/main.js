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
  
  // Create enhanced Financial Tools menu with submenus
  ui.createMenu('Financial Tools')
    .addItem('Generate Overview', 'createFinancialOverview')
    .addSeparator()
    .addSubMenu(ui.createMenu('Reports')
      .addItem('Monthly Spending Report', 'generateMonthlySpendingReport')
      .addItem('Yearly Summary', 'generateYearlySummary')
      .addItem('Category Breakdown', 'generateCategoryBreakdown')
      .addItem('Savings Analysis', 'generateSavingsAnalysis'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Visualizations')
      .addItem('Spending Trends Chart', 'createSpendingTrendsChart')
      .addItem('Budget vs Actual', 'createBudgetVsActualChart')
      .addItem('Income vs Expenses', 'createIncomeVsExpensesChart')
      .addItem('Category Pie Chart', 'createCategoryPieChart'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Financial Analysis')
      .addItem('Suggest Savings Opportunities', 'suggestSavingsOpportunities')
      .addItem('Spending Anomaly Detection', 'detectSpendingAnomalies')
      .addItem('Fixed vs Variable Expenses', 'analyzeFixedVsVariableExpenses')
      .addItem('Cash Flow Forecast', 'generateCashFlowForecast'))
    .addSeparator()
    .addItem('Set Budget Targets', 'setBudgetTargets')
    .addItem('Setup Email Reports', 'setupEmailReports')
    .addToUi();
}
