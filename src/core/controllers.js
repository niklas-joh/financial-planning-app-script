/**
 * @fileoverview Controllers Module for Financial Planning Tools.
 * Centralizes UI-triggered actions and coordinates between UI and services.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const ControllersModule = (function() {
  /**
   * Constructor for the ControllersModule.
   * @param {object} configInstance - An instance of ConfigModule.
   * @param {object} uiServiceInstance - An instance of UIServiceModule.
   * @param {object} errorServiceInstance - An instance of ErrorServiceModule.
   * @constructor
   */
  function ControllersModuleConstructor(configInstance, uiServiceInstance, errorServiceInstance) {
    this.config = configInstance;
    this.uiService = uiServiceInstance;
    this.errorService = errorServiceInstance;

    // Initialize wrapped methods after construction
    this._initializeWrappedMethods();
  }

  ControllersModuleConstructor.prototype._wrapWithFeedback = function(fn, startMessage, successMessage, errorMessage) {
    const self = this;
    return function(...args) {
      try {
        if (startMessage) {
          self.uiService.showLoadingSpinner(startMessage);
        }
        // Bind fn to self to ensure it has access to this.config, etc. if it's a method of ControllersModule
        const result = fn.apply(self, args);
        self.uiService.hideLoadingSpinner();
        if (successMessage) {
          self.uiService.showSuccessNotification(successMessage);
        }
        return result;
      } catch (error) {
        self.uiService.hideLoadingSpinner();
        self.errorService.handle(
          error,
          errorMessage || 'An error occurred while performing the operation.'
        );
        throw error; // Re-throw so it can be caught by GAS if needed
      }
    };
  };

  // Core logic for controller actions (unwrapped)
  // These will be wrapped and assigned to the instance in _initializeWrappedMethods
  const coreLogic = {
    createFinancialOverview: function() {
      return FinancialPlanner.FinanceOverview.create();
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
      const overviewSheetName = this.config.getSheetNames().OVERVIEW;
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
    // Add other core logic functions here...
    suggestSavingsOpportunities: function() { /* Placeholder */ console.log("Suggesting savings..."); },
    detectSpendingAnomalies: function() { /* Placeholder */ console.log("Detecting anomalies..."); },
    analyzeFixedVsVariableExpenses: function() { /* Placeholder */ console.log("Analyzing fixed vs variable..."); },
    generateCashFlowForecast: function() { /* Placeholder */ console.log("Generating cash flow forecast..."); },
    setBudgetTargets: function() { /* Placeholder */ console.log("Setting budget targets..."); },
    setupEmailReports: function() { /* Placeholder */ console.log("Setting up email reports..."); },
  };

  // Method to initialize wrapped versions of controller actions
  ControllersModuleConstructor.prototype._initializeWrappedMethods = function() {
    this.createFinancialOverview_Wrapped = this._wrapWithFeedback(coreLogic.createFinancialOverview, 'Generating financial overview...', 'Financial overview generated successfully!', 'Failed to generate financial overview');
    this.generateMonthlySpendingReport_Wrapped = this._wrapWithFeedback(coreLogic.generateMonthlySpendingReport, 'Generating monthly spending report...', 'Monthly spending report generated successfully!', 'Failed to generate monthly spending report');
    this.showKeyMetrics_Wrapped = this._wrapWithFeedback(coreLogic.showKeyMetrics, 'Analyzing financial data...', 'Key metrics displayed successfully!', 'Failed to display key metrics');
    this.generateYearlySummary_Wrapped = this._wrapWithFeedback(coreLogic.generateYearlySummary, 'Generating yearly summary report...', 'Yearly summary report generated successfully!', 'Failed to generate yearly summary report');
    this.generateCategoryBreakdown_Wrapped = this._wrapWithFeedback(coreLogic.generateCategoryBreakdown, 'Generating category breakdown report...', 'Category breakdown report generated successfully!', 'Failed to generate category breakdown report');
    this.generateSavingsAnalysis_Wrapped = this._wrapWithFeedback(coreLogic.generateSavingsAnalysis, 'Generating savings analysis report...', 'Savings analysis report generated successfully!', 'Failed to generate savings analysis report');
    this.createSpendingTrendsChart_Wrapped = this._wrapWithFeedback(coreLogic.createSpendingTrendsChart, 'Creating spending trends chart...', 'Spending trends chart created successfully!', 'Failed to create spending trends chart');
    this.createBudgetVsActualChart_Wrapped = this._wrapWithFeedback(coreLogic.createBudgetVsActualChart, 'Creating budget vs actual chart...', 'Budget vs actual chart created successfully!', 'Failed to create budget vs actual chart');
    this.createIncomeVsExpensesChart_Wrapped = this._wrapWithFeedback(coreLogic.createIncomeVsExpensesChart, 'Creating income vs expenses chart...', 'Income vs expenses chart created successfully!', 'Failed to create income vs expenses chart');
    this.createCategoryPieChart_Wrapped = this._wrapWithFeedback(coreLogic.createCategoryPieChart, 'Creating category pie chart...', 'Category pie chart created successfully!', 'Failed to create category pie chart');
    this.toggleShowSubCategories_Wrapped = this._wrapWithFeedback(coreLogic.toggleShowSubCategories, 'Updating display preferences...', 'Display preferences updated successfully!', 'Failed to update display preferences');
    this.refreshCache_Wrapped = this._wrapWithFeedback(coreLogic.refreshCache, 'Refreshing all caches...', 'Caches refreshed successfully!', 'Failed to refresh one or more caches');
    
    // Placeholders for "Coming Soon" features
    this.suggestSavingsOpportunities_Wrapped = this._wrapWithFeedback(coreLogic.suggestSavingsOpportunities, 'Working...', 'Coming soon!', 'Operation failed');
    this.detectSpendingAnomalies_Wrapped = this._wrapWithFeedback(coreLogic.detectSpendingAnomalies, 'Working...', 'Coming soon!', 'Operation failed');
    this.analyzeFixedVsVariableExpenses_Wrapped = this._wrapWithFeedback(coreLogic.analyzeFixedVsVariableExpenses, 'Working...', 'Coming soon!', 'Operation failed');
    this.generateCashFlowForecast_Wrapped = this._wrapWithFeedback(coreLogic.generateCashFlowForecast, 'Working...', 'Coming soon!', 'Operation failed');
    this.setBudgetTargets_Wrapped = this._wrapWithFeedback(coreLogic.setBudgetTargets, 'Working...', 'Coming soon!', 'Operation failed');
    this.setupEmailReports_Wrapped = this._wrapWithFeedback(coreLogic.setupEmailReports, 'Working...', 'Coming soon!', 'Operation failed');
  };

  // Event Handlers (not wrapped with UI feedback by default, they handle errors internally)
  ControllersModuleConstructor.prototype.onOpen = function() {
    try {
      const ui = SpreadsheetApp.getUi();
      // Menu items will call global functions, which in turn call the _Wrapped methods on the instance
      ui.createMenu('📊 Financial Tools')
        .addItem('📈 Generate Overview', 'createFinancialOverview_Global')
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
      // Use the injected error service instance
      if (this.errorService && typeof this.errorService.log === 'function') {
        this.errorService.log(this.errorService.create("Failed to create menu in onOpen", {originalError: error, severity: 'high'}));
      } else {
        console.error("Failed to create menu (ErrorService not available):", error.message, error.stack);
      }
    }
  };

  ControllersModuleConstructor.prototype.onEdit = function(e) {
    try {
      const sheet = e.range.getSheet();
      const sheetName = sheet.getName();
      const overviewSheetName = this.config.getSheetNames().OVERVIEW;
      const transactionsSheetName = this.config.getSheetNames().TRANSACTIONS;

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
      if (this.errorService && typeof this.errorService.log === 'function') {
         this.errorService.log(this.errorService.create("Error handling onEdit event", {originalError: error, eventDetails: e ? JSON.stringify(e) : 'N/A', severity: 'medium'}));
      } else {
        console.error("Error handling edit event (ErrorService not available):", error.message, error.stack);
      }
    }
  };

  return ControllersModuleConstructor;
})();

// --- Global Functions for Apps Script Triggers & Menu Items ---
// These functions will call methods on the FinancialPlanner.Controllers INSTANCE.

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
  } else {
    // console.warn('FinancialPlanner.Controllers.onEdit not available at onEdit trigger.');
  }
}

// Helper to create global functions that call the wrapped instance methods
function createGlobalControllerAction(methodName) {
  // eslint-disable-next-line no-unused-vars
  this[methodName + '_Global'] = function(...args) { // Use 'this' to attach to global scope in Apps Script
    if (FinancialPlanner && FinancialPlanner.Controllers && typeof FinancialPlanner.Controllers[methodName + '_Wrapped'] === 'function') {
      try {
        return FinancialPlanner.Controllers[methodName + '_Wrapped'](...args);
      } catch (e) {
        // Error already handled by _wrapWithFeedback, but good to log it happened at global call level too.
        // FinancialPlanner.ErrorService might not be available here if Controllers itself failed to init.
        console.error(`Error during global call to ${methodName}_Wrapped: ${e.message}`);
        // Optionally, show a generic UI error if possible and not redundant
        // SpreadsheetApp.getUi().alert("An unexpected error occurred. Please check logs.");
      }
    } else {
      const msg = `Controller action ${methodName}_Wrapped not available.`;
      console.error(msg);
      if (FinancialPlanner && FinancialPlanner.Controllers && FinancialPlanner.Controllers.uiService) {
         FinancialPlanner.Controllers.uiService.showErrorNotification('Error', msg);
      } else {
        SpreadsheetApp.getUi().alert(msg);
      }
    }
  };
}

// Create global functions for all wrapped controller actions
// These names must match what's used in the onOpen menu creation.
createGlobalControllerAction('createFinancialOverview');
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
