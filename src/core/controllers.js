/**
 * @fileoverview Controllers Module for Financial Planning Tools.
 * Centralizes UI-triggered actions and coordinates between UI and services.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

/**
 * @module core/controllers
 * @description IIFE for ControllersModule. Encapsulates UI-triggered actions and coordinates services.
 * It provides methods that are typically called from custom menu items or event triggers (onOpen, onEdit).
 * These methods often wrap core service logic with UI feedback (loading spinners, notifications).
 */
// eslint-disable-next-line no-unused-vars
const ControllersModule = (function() {
  /**
   * Constructor for the ControllersModule.
   * Initializes the controller with necessary service instances.
   * @constructor
   * @param {ConfigModule} configInstance - An instance of the ConfigModule for accessing configuration.
   * @param {UIServiceModule} uiServiceInstance - An instance of the UIServiceModule for UI interactions.
   * @param {ErrorServiceModule} errorServiceInstance - An instance of the ErrorServiceModule for error handling.
   */
  function ControllersModuleConstructor(configInstance, uiServiceInstance, errorServiceInstance) {
    this.config = configInstance;
    this.uiService = uiServiceInstance;
    this.errorService = errorServiceInstance;

    // Initialize wrapped methods after construction
    this._initializeWrappedMethods();
  }

  /**
   * Wraps a given function with UI feedback (loading spinner, success/error notifications).
   * This is a higher-order function used to standardize UI feedback for controller actions.
   * @private
   * @memberof ControllersModuleConstructor
   * @instance
   * @param {function} fn - The function to wrap. This function will be applied with the controller instance as `this`.
   * @param {string} [startMessage] - Optional message to display in a loading spinner before executing the function.
   * @param {string} [successMessage] - Optional message for a success notification if the function executes without errors.
   * @param {string} [errorMessage] - Optional custom error message for the notification if the function throws an error.
   * @returns {function} A new function that, when called, executes the original function with UI feedback.
   * @throws {Error} Re-throws any error caught from the wrapped function after handling UI feedback and logging.
   */
  ControllersModuleConstructor.prototype._wrapWithFeedback = function(fn, startMessage, successMessage, errorMessage) {
    const self = this; // Preserve 'this' context of the ControllersModule instance
    return function(...args) {
      try {
        if (startMessage) {
          self.uiService.showLoadingSpinner(startMessage);
        }
        // Bind fn to self to ensure it has access to this.config, this.uiService, etc.
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
        throw error; // Re-throw so it can be caught by GAS or other callers if needed
      }
    };
  };

  // Core logic for controller actions (unwrapped)
  // These will be wrapped and assigned to the instance in _initializeWrappedMethods
  /**
   * @private
   * @const {object} coreLogic
   * @description An object containing the core (unwrapped) logic for controller actions.
   * These functions are later wrapped with UI feedback by `_initializeWrappedMethods`.
   * Each method typically calls a corresponding service method from the `FinancialPlanner` global namespace.
   */
  const coreLogic = {
    /** Calls FinanceOverview service to create/update the overview sheet. */
    createFinancialOverview: function() {
      return FinancialPlanner.FinanceOverview.create();
    },
    /** Calls MonthlySpendingReport service to generate the report. */
    generateMonthlySpendingReport: function() {
      return FinancialPlanner.MonthlySpendingReport.generate();
    },
    /** Calls FinancialAnalysisService to display key financial metrics. */
    showKeyMetrics: function() {
      return FinancialPlanner.FinancialAnalysisService.showKeyMetrics();
    },
    /** Calls ReportService to generate a yearly summary. */
    generateYearlySummary: function() {
      return FinancialPlanner.ReportService.generateYearlySummary();
    },
    /** Calls ReportService to generate a category breakdown. */
    generateCategoryBreakdown: function() {
      return FinancialPlanner.ReportService.generateCategoryBreakdown();
    },
    /** Calls ReportService to generate a savings analysis. */
    generateSavingsAnalysis: function() {
      return FinancialPlanner.ReportService.generateSavingsAnalysis();
    },
    /** Calls VisualizationService to create a spending trends chart. */
    createSpendingTrendsChart: function() {
      return FinancialPlanner.VisualizationService.createSpendingTrendsChart();
    },
    /** Calls VisualizationService to create a budget vs. actual chart. */
    createBudgetVsActualChart: function() {
      return FinancialPlanner.VisualizationService.createBudgetVsActualChart();
    },
    /** Calls VisualizationService to create an income vs. expenses chart. */
    createIncomeVsExpensesChart: function() {
      return FinancialPlanner.VisualizationService.createIncomeVsExpensesChart();
    },
    /** Calls VisualizationService to create a category pie chart. */
    createCategoryPieChart: function() {
      return FinancialPlanner.VisualizationService.createCategoryPieChart();
    },
    /** 
     * Toggles the visibility of sub-category columns in the overview sheet.
     * Uses SettingsService to persist the preference and directly manipulates sheet columns.
     * @returns {boolean} The new state of sub-category visibility (true if shown, false if hidden).
     */
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
    /** 
     * Refreshes application caches.
     * Invalidates all general caches via CacheService and refreshes DropdownService cache.
     * @returns {boolean} Always returns true upon completion.
     */
    refreshCache: function() {
      if (FinancialPlanner.CacheService && FinancialPlanner.CacheService.invalidateAll) {
        FinancialPlanner.CacheService.invalidateAll();
      }
      if (FinancialPlanner.DropdownService && FinancialPlanner.DropdownService.refreshCache) {
        FinancialPlanner.DropdownService.refreshCache();
      }
      return true;
    },
    // --- Placeholder/Coming Soon Features ---
    /** Placeholder for suggesting savings opportunities. */
    suggestSavingsOpportunities: function() { /* Placeholder */ console.log("Suggesting savings..."); },
    /** Placeholder for detecting spending anomalies. */
    detectSpendingAnomalies: function() { /* Placeholder */ console.log("Detecting anomalies..."); },
    /** Placeholder for analyzing fixed vs. variable expenses. */
    analyzeFixedVsVariableExpenses: function() { /* Placeholder */ console.log("Analyzing fixed vs variable..."); },
    /** Placeholder for generating a cash flow forecast. */
    generateCashFlowForecast: function() { /* Placeholder */ console.log("Generating cash flow forecast..."); },
    /** Placeholder for setting budget targets. */
    setBudgetTargets: function() { /* Placeholder */ console.log("Setting budget targets..."); },
    /** Placeholder for setting up email reports. */
    setupEmailReports: function() { /* Placeholder */ console.log("Setting up email reports..."); },
  };

  /**
   * Initializes wrapped versions of all methods defined in `coreLogic`.
   * Each core logic function is wrapped with UI feedback (spinner, notifications)
   * and assigned as a new method on the `ControllersModule` instance (e.g., `this.createFinancialOverview_Wrapped`).
   * This method is called by the constructor.
   * @private
   * @memberof ControllersModuleConstructor
   * @instance
   */
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
  /**
   * Handles the `onOpen` simple trigger for the Google Apps Script project.
   * Creates the custom "Financial Tools" menu in the spreadsheet UI.
   * Errors during menu creation are logged via the ErrorService.
   * @memberof ControllersModuleConstructor
   * @instance
   */
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

  /**
   * Handles the `onEdit` simple trigger for the Google Apps Script project.
   * Delegates edit handling to specific services based on the edited sheet.
   * For example, edits in the 'Overview' sheet might be handled by `FinanceOverview.handleEdit`,
   * and edits in 'Transactions' by `DropdownService.handleEdit`.
   * Errors during edit handling are logged via the ErrorService.
   * @memberof ControllersModuleConstructor
   * @instance
   * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object passed by the `onEdit` trigger.
   */
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

/**
 * Global `onOpen` simple trigger function for Google Apps Script.
 * This function is automatically executed when the spreadsheet is opened.
 * It delegates to the `onOpen` method of the instantiated `FinancialPlanner.Controllers`.
 * If `FinancialPlanner.Controllers` is not available, it logs an error and adds an error menu item.
 * @global
 */
function onOpen() {
  if (FinancialPlanner && FinancialPlanner.Controllers && typeof FinancialPlanner.Controllers.onOpen === 'function') {
    FinancialPlanner.Controllers.onOpen();
  } else {
    console.error('FinancialPlanner.Controllers or FinancialPlanner.Controllers.onOpen not available at onOpen trigger.');
    SpreadsheetApp.getUi().createMenu('⚠️ Error').addItem('Setup Incomplete. Check Logs.', 'function(){};').addToUi();
  }
}

/**
 * Global `onEdit` simple trigger function for Google Apps Script.
 * This function is automatically executed when a user edits a cell in the spreadsheet.
 * It delegates to the `onEdit` method of the instantiated `FinancialPlanner.Controllers`.
 * @global
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object passed by the `onEdit` trigger.
 */
function onEdit(e) {
  if (FinancialPlanner && FinancialPlanner.Controllers && typeof FinancialPlanner.Controllers.onEdit === 'function') {
    FinancialPlanner.Controllers.onEdit(e);
  } else {
    // console.warn('FinancialPlanner.Controllers.onEdit not available at onEdit trigger.'); // Kept for debugging if needed
  }
}

/**
 * Helper function to create global functions that call wrapped methods on the `FinancialPlanner.Controllers` instance.
 * For each `methodName` provided, it creates a new global function named `${methodName}_Global`.
 * This global function, when called (e.g., from a menu item), will invoke the corresponding
 * `${methodName}_Wrapped` method on the `FinancialPlanner.Controllers` instance.
 * It includes error handling for cases where the controller or wrapped method might not be available.
 * @global
 * @param {string} methodName - The base name of the controller method to create a global wrapper for (e.g., 'createFinancialOverview').
 */
function createGlobalControllerAction(methodName) {
  // eslint-disable-next-line no-unused-vars
  this[methodName + '_Global'] = function(...args) { // Use 'this' to attach to global scope in Apps Script
    if (FinancialPlanner && FinancialPlanner.Controllers && typeof FinancialPlanner.Controllers[methodName + '_Wrapped'] === 'function') {
      try {
        return FinancialPlanner.Controllers[methodName + '_Wrapped'](...args);
      } catch (e) {
        // Error is typically already handled by _wrapWithFeedback, but this logs it at the global call level too.
        // FinancialPlanner.ErrorService might not be available here if Controllers itself failed to initialize.
        console.error(`Error during global call to ${methodName}_Wrapped: ${e.message}`);
        // Optionally, show a generic UI error if possible and not redundant with _wrapWithFeedback's handling.
        // SpreadsheetApp.getUi().alert("An unexpected error occurred. Please check logs.");
      }
    } else {
      const msg = `Controller action '${methodName}_Wrapped' not available. Check if FinancialPlanner.Controllers is initialized.`;
      console.error(msg);
      // Attempt to use the UI service from the potentially available controller instance for a user-facing error.
      if (FinancialPlanner && FinancialPlanner.Controllers && FinancialPlanner.Controllers.uiService && typeof FinancialPlanner.Controllers.uiService.showErrorNotification === 'function') {
         FinancialPlanner.Controllers.uiService.showErrorNotification('Action Failed', msg);
      } else {
        // Fallback to a simple alert if the UI service isn't accessible.
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
