/**
 * @fileoverview Controllers Module for Financial Planning Tools.
 * Centralizes UI-triggered actions and coordinates between UI and services.
 * @module core/controllers
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Application version information
 * @const
 */
FinancialPlanner.VERSION = '3.0.0';

/**
 * Application metadata
 * @const
 */
FinancialPlanner.META = {
  name: 'Financial Planning Tools',
  description: 'Google Apps Script project for financial planning and analysis',
  author: 'Financial Planning Team',
  lastUpdated: '2025-05-11'
};

/**
 * Controllers - Coordinates UI actions and service calls.
 * @namespace FinancialPlanner.Controllers
 */
FinancialPlanner.Controllers = (function() {
  /**
   * Wraps a given function with UI feedback (loading spinner, success/error notifications).
   * @private
   * @param {function} fn - The function to wrap.
   * @param {string} [startMessage] - Optional message to display before executing.
   * @param {string} [successMessage] - Optional success message.
   * @param {string} [errorMessage] - Optional custom error message.
   * @returns {function} A new function with UI feedback.
   */
  function wrapWithFeedback(fn, startMessage, successMessage, errorMessage) {
    return function() {
      try {
        if (startMessage) {
          FinancialPlanner.UIService.showLoadingSpinner(startMessage);
        }
        const result = fn.apply(this, arguments);
        FinancialPlanner.UIService.hideLoadingSpinner();
        if (successMessage) {
          FinancialPlanner.UIService.showSuccessNotification(successMessage);
        }
        return result;
      } catch (error) {
        FinancialPlanner.UIService.hideLoadingSpinner();
        FinancialPlanner.ErrorService.handle(
          error,
          errorMessage || 'An error occurred while performing the operation.'
        );
        throw error;
      }
    };
  }

  // Core logic for controller actions
  const coreLogic = {
    createFinancialOverview: function() {
      return FinancialPlanner.FinanceOverview.create();
    },
    connectBankAccount: function() {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('services/plaid-link')
        .setWidth(600)
        .setHeight(500);
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Connect Bank Account');
    },
    importTransactions: function() {
      const endDate = new Date();
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - 30);
      
      const dateFormat = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const endFormat = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      const result = FinancialPlanner.PlaidService.getTransactions(dateFormat, endFormat);
      const count = FinancialPlanner.PlaidService.importToSheet(result.transactions);
      
      return count;
    },
    generateMonthlySpendingReport: function() {
      return FinancialPlanner.MonthlySpendingReport.generate();
    },
    showKeyMetrics: function() {
      return FinancialPlanner.FinancialAnalysisService.showKeyMetrics();
    },
    generateYearlySummary: function() {
      return FinancialPlanner.ReportService.generateYearlySummary();
    },
    generateCategoryBreakdown: function() {
      return FinancialPlanner.ReportService.generateCategoryBreakdown();
    },
    generateSavingsAnalysis: function() {
      return FinancialPlanner.ReportService.generateSavingsAnalysis();
    },
    createSpendingTrendsChart: function() {
      return FinancialPlanner.VisualizationService.createSpendingTrendsChart();
    },
    createBudgetVsActualChart: function() {
      return FinancialPlanner.VisualizationService.createBudgetVsActualChart();
    },
    createIncomeVsExpensesChart: function() {
      return FinancialPlanner.VisualizationService.createIncomeVsExpensesChart();
    },
    createCategoryPieChart: function() {
      return FinancialPlanner.VisualizationService.createCategoryPieChart();
    },
    toggleShowSubCategories: function() {
      const newValue = FinancialPlanner.SettingsService.toggleShowSubCategories();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const overviewSheetName = FinancialPlanner.Config.getSheetNames().OVERVIEW;
      const overviewSheet = ss.getSheetByName(overviewSheetName);
      if (overviewSheet) {
        if (newValue) {
          overviewSheet.showColumns(3, 1);
        } else {
          overviewSheet.hideColumns(3, 1);
        }
      }
      return newValue;
    },
    refreshCache: function() {
      if (FinancialPlanner.CacheService && FinancialPlanner.CacheService.invalidateAll) {
        FinancialPlanner.CacheService.invalidateAll();
      }
      if (FinancialPlanner.DropdownService && FinancialPlanner.DropdownService.refreshCache) {
        FinancialPlanner.DropdownService.refreshCache();
      }
      return true;
    },
    suggestSavingsOpportunities: function() { console.log("Suggesting savings..."); },
    detectSpendingAnomalies: function() { console.log("Detecting anomalies..."); },
    analyzeFixedVsVariableExpenses: function() { console.log("Analyzing fixed vs variable..."); },
    generateCashFlowForecast: function() { console.log("Generating cash flow forecast..."); },
    setBudgetTargets: function() { console.log("Setting budget targets..."); },
    setupEmailReports: function() { console.log("Setting up email reports..."); }
  };

  // Public API
  return {
    // Wrapped methods
    createFinancialOverview_Wrapped: wrapWithFeedback(coreLogic.createFinancialOverview, 'Generating financial overview...', 'Financial overview generated successfully!', 'Failed to generate financial overview'),
    connectBankAccount_Wrapped: wrapWithFeedback(coreLogic.connectBankAccount, null, null, 'Failed to open bank connection dialog'),
    importTransactions_Wrapped: wrapWithFeedback(coreLogic.importTransactions, 'Importing transactions from bank...', 'Transactions imported successfully!', 'Failed to import transactions'),
    generateMonthlySpendingReport_Wrapped: wrapWithFeedback(coreLogic.generateMonthlySpendingReport, 'Generating monthly spending report...', 'Monthly spending report generated successfully!', 'Failed to generate monthly spending report'),
    showKeyMetrics_Wrapped: wrapWithFeedback(coreLogic.showKeyMetrics, 'Analyzing financial data...', 'Key metrics displayed successfully!', 'Failed to display key metrics'),
    generateYearlySummary_Wrapped: wrapWithFeedback(coreLogic.generateYearlySummary, 'Generating yearly summary report...', 'Yearly summary report generated successfully!', 'Failed to generate yearly summary report'),
    generateCategoryBreakdown_Wrapped: wrapWithFeedback(coreLogic.generateCategoryBreakdown, 'Generating category breakdown report...', 'Category breakdown report generated successfully!', 'Failed to generate category breakdown report'),
    generateSavingsAnalysis_Wrapped: wrapWithFeedback(coreLogic.generateSavingsAnalysis, 'Generating savings analysis report...', 'Savings analysis report generated successfully!', 'Failed to generate savings analysis report'),
    createSpendingTrendsChart_Wrapped: wrapWithFeedback(coreLogic.createSpendingTrendsChart, 'Creating spending trends chart...', 'Spending trends chart created successfully!', 'Failed to create spending trends chart'),
    createBudgetVsActualChart_Wrapped: wrapWithFeedback(coreLogic.createBudgetVsActualChart, 'Creating budget vs actual chart...', 'Budget vs actual chart created successfully!', 'Failed to create budget vs actual chart'),
    createIncomeVsExpensesChart_Wrapped: wrapWithFeedback(coreLogic.createIncomeVsExpensesChart, 'Creating income vs expenses chart...', 'Income vs expenses chart created successfully!', 'Failed to create income vs expenses chart'),
    createCategoryPieChart_Wrapped: wrapWithFeedback(coreLogic.createCategoryPieChart, 'Creating category pie chart...', 'Category pie chart created successfully!', 'Failed to create category pie chart'),
    toggleShowSubCategories_Wrapped: wrapWithFeedback(coreLogic.toggleShowSubCategories, 'Updating display preferences...', 'Display preferences updated successfully!', 'Failed to update display preferences'),
    refreshCache_Wrapped: wrapWithFeedback(coreLogic.refreshCache, 'Refreshing all caches...', 'Caches refreshed successfully!', 'Failed to refresh one or more caches'),
    suggestSavingsOpportunities_Wrapped: wrapWithFeedback(coreLogic.suggestSavingsOpportunities, 'Working...', 'Coming soon!', 'Operation failed'),
    detectSpendingAnomalies_Wrapped: wrapWithFeedback(coreLogic.detectSpendingAnomalies, 'Working...', 'Coming soon!', 'Operation failed'),
    analyzeFixedVsVariableExpenses_Wrapped: wrapWithFeedback(coreLogic.analyzeFixedVsVariableExpenses, 'Working...', 'Coming soon!', 'Operation failed'),
    generateCashFlowForecast_Wrapped: wrapWithFeedback(coreLogic.generateCashFlowForecast, 'Working...', 'Coming soon!', 'Operation failed'),
    setBudgetTargets_Wrapped: wrapWithFeedback(coreLogic.setBudgetTargets, 'Working...', 'Coming soon!', 'Operation failed'),
    setupEmailReports_Wrapped: wrapWithFeedback(coreLogic.setupEmailReports, 'Working...', 'Coming soon!', 'Operation failed'),

    // Event Handlers
    onOpen: function() {
      try {
        const ui = SpreadsheetApp.getUi();
        ui.createMenu('📊 Financial Tools')
          .addItem('📈 Generate Overview', 'createFinancialOverview_Global')
          .addSeparator()
          .addSubMenu(ui.createMenu('🏦 Bank Integration')
            .addItem('🔗 Connect Bank Account', 'connectBankAccount_Global')
            .addItem('📥 Import Transactions', 'importTransactions_Global'))
          .addSeparator()
          .addSubMenu(ui.createMenu('📋 Reports')
            .addItem('📝 Monthly Spending Report', 'generateMonthlySpendingReport_Global')
            .addItem('📅 Yearly Summary', 'generateYearlySummary_Global')
            .addItem('🔍 Category Breakdown', 'generateCategoryBreakdown_Global')
            .addItem('💰 Savings Analysis', 'generateSavingsAnalysis_Global'))
          .addSeparator()
          .addSubMenu(ui.createMenu('📊 Visualizations')
            .addItem('📉 Spending Trends Chart', 'createSpendingTrendsChart_Global')
            .addItem('⚖️ Budget vs Actual', 'createBudgetVsActualChart_Global')
            .addItem('💹 Income vs Expenses', 'createIncomeVsExpensesChart_Global')
            .addItem('🍩 Category Pie Chart', 'createCategoryPieChart_Global'))
          .addSeparator()
          .addSubMenu(ui.createMenu('🧮 Financial Analysis')
            .addItem('📊 Key Metrics', 'showKeyMetrics_Global')
            .addItem('💡 Suggest Savings (Soon)', 'suggestSavingsOpportunities_Global')
            .addItem('⚠️ Anomalies (Soon)', 'detectSpendingAnomalies_Global')
            .addItem('📌 Fixed/Variable (Soon)', 'analyzeFixedVsVariableExpenses_Global')
            .addItem('🔮 Cash Flow (Soon)', 'generateCashFlowForecast_Global'))
          .addSeparator()
          .addSubMenu(ui.createMenu('⚙️ Settings')
            .addItem('🔄 Toggle Sub-Categories', 'toggleShowSubCategories_Global')
            .addItem('🎯 Set Budgets (Soon)', 'setBudgetTargets_Global')
            .addItem('📧 Email Reports (Soon)', 'setupEmailReports_Global')
            .addItem('🔄 Refresh Cache', 'refreshCache_Global'))
          .addToUi();
      } catch (error) {
        if (FinancialPlanner.ErrorService && typeof FinancialPlanner.ErrorService.log === 'function') {
          FinancialPlanner.ErrorService.log(FinancialPlanner.ErrorService.create("Failed to create menu in onOpen", {originalError: error, severity: 'high'}));
        } else {
          console.error("Failed to create menu (ErrorService not available):", error.message, error.stack);
        }
      }
    },

    onEdit: function(e) {
      try {
        const sheet = e.range.getSheet();
        const sheetName = sheet.getName();
        const overviewSheetName = FinancialPlanner.Config.getSheetNames().OVERVIEW;
        const transactionsSheetName = FinancialPlanner.Config.getSheetNames().TRANSACTIONS;

        if (sheetName === overviewSheetName) {
          if (FinancialPlanner.FinanceOverview && FinancialPlanner.FinanceOverview.handleEdit) {
            FinancialPlanner.FinanceOverview.handleEdit(e);
          }
        } else if (sheetName === transactionsSheetName) {
          if (FinancialPlanner.DropdownService && FinancialPlanner.DropdownService.handleEdit) {
            FinancialPlanner.DropdownService.handleEdit(e);
          }
        }
      } catch (error) {
        if (FinancialPlanner.ErrorService && typeof FinancialPlanner.ErrorService.log === 'function') {
          FinancialPlanner.ErrorService.log(FinancialPlanner.ErrorService.create("Error handling onEdit event", {originalError: error, eventDetails: e ? JSON.stringify(e) : 'N/A', severity: 'medium'}));
        } else {
          console.error("Error handling edit event (ErrorService not available):", error.message, error.stack);
        }
      }
    }
  };
})();

// Global trigger functions
function onOpen() {
  if (FinancialPlanner && FinancialPlanner.Controllers && typeof FinancialPlanner.Controllers.onOpen === 'function') {
    FinancialPlanner.Controllers.onOpen();
  } else {
    console.error('FinancialPlanner.Controllers or FinancialPlanner.Controllers.onOpen not available at onOpen trigger.');
    SpreadsheetApp.getUi().createMenu('⚠️ Error').addItem('Setup Incomplete. Check Logs.', 'function(){};').addToUi();
  }
}

function onEdit(e) {
  if (FinancialPlanner && FinancialPlanner.Controllers && typeof FinancialPlanner.Controllers.onEdit === 'function') {
    FinancialPlanner.Controllers.onEdit(e);
  }
}

// Global action functions
function createGlobalControllerAction(methodName) {
  this[methodName + '_Global'] = function() {
    if (FinancialPlanner && FinancialPlanner.Controllers && typeof FinancialPlanner.Controllers[methodName + '_Wrapped'] === 'function') {
      try {
        return FinancialPlanner.Controllers[methodName + '_Wrapped'].apply(FinancialPlanner.Controllers, arguments);
      } catch (e) {
        console.error('Error during global call to ' + methodName + '_Wrapped: ' + e.message);
      }
    } else {
      const msg = 'Controller action ' + methodName + '_Wrapped not available. Check if FinancialPlanner.Controllers is initialized.';
      console.error(msg);
      if (FinancialPlanner && FinancialPlanner.Controllers && FinancialPlanner.Controllers.uiService && typeof FinancialPlanner.UIService.showErrorNotification === 'function') {
        FinancialPlanner.UIService.showErrorNotification('Action Failed', msg);
      } else {
        SpreadsheetApp.getUi().alert(msg);
      }
    }
  };
}

// Create global functions for all wrapped controller actions
createGlobalControllerAction('createFinancialOverview');
createGlobalControllerAction('connectBankAccount');
createGlobalControllerAction('importTransactions');
createGlobalControllerAction('generateMonthlySpendingReport');
createGlobalControllerAction('showKeyMetrics');
createGlobalControllerAction('generateYearlySummary');
createGlobalControllerAction('generateCategoryBreakdown');
createGlobalControllerAction('generateSavingsAnalysis');
createGlobalControllerAction('createSpendingTrendsChart');
createGlobalControllerAction('createBudgetVsActualChart');
createGlobalControllerAction('createIncomeVsExpensesChart');
createGlobalControllerAction('createCategoryPieChart');
createGlobalControllerAction('toggleShowSubCategories');
createGlobalControllerAction('refreshCache');
createGlobalControllerAction('suggestSavingsOpportunities');
createGlobalControllerAction('detectSpendingAnomalies');
createGlobalControllerAction('analyzeFixedVsVariableExpenses');
createGlobalControllerAction('generateCashFlowForecast');
createGlobalControllerAction('setBudgetTargets');
createGlobalControllerAction('setupEmailReports');

// Log successful initialization
Logger.log('FinancialPlanner modules loaded successfully. Version: ' + FinancialPlanner.VERSION);
