/**
 * Financial Planning Tools - Controllers
 * 
 * This file provides a centralized set of controller functions that serve as
 * entry points for UI-triggered actions. These functions coordinate between
 * the UI and the underlying services.
 */

// Create the Controllers module within the FinancialPlanner namespace
/**
 * @namespace FinancialPlanner.Controllers
 * @param {FinancialPlanner.Config} config - The configuration service.
 * @param {FinancialPlanner.UIService} uiService - The UI service.
 * @param {FinancialPlanner.ErrorService} errorService - The error handling service.
 */
FinancialPlanner.Controllers = (function(config, uiService, errorService) {
  // Private variables and functions
  
  /**
   * Wraps a given function with UI feedback (loading, success, error messages)
   * and standardized error handling using `ErrorService`.
   *
   * @param {function(...any): any} fn - The function to wrap. This function will be called with the original arguments.
   * @param {string} [startMessage] - Optional message to display via `uiService.showLoadingSpinner` before executing `fn`.
   * @param {string} [successMessage] - Optional message to display via `uiService.showSuccessNotification` after `fn` executes successfully.
   * @param {string} [errorMessage] - Optional user-friendly message to pass to `errorService.handle` if `fn` throws an error.
   *                                  Defaults to "An error occurred while performing the operation.".
   * @return {function(...any): any} The wrapped function. This function will return the result of `fn` or re-throw the error after handling.
   * @throws {Error} Re-throws the error caught from the execution of `fn` after it has been handled by `ErrorService`.
   * @private
   */
  function wrapWithFeedback(fn, startMessage, successMessage, errorMessage) {
    return function() {
      try {
        // Show loading message if provided
        if (startMessage) {
          uiService.showLoadingSpinner(startMessage);
        }
        
        // Call the original function
        const result = fn.apply(this, arguments);
        
        // Hide loading spinner
        uiService.hideLoadingSpinner();
        
        // Show success message if provided
        if (successMessage) {
          uiService.showSuccessNotification(successMessage);
        }
        
        return result;
      } catch (error) {
        // Hide loading spinner
        uiService.hideLoadingSpinner();
        
        // Handle the error
        errorService.handle(
          error,
          errorMessage || "An error occurred while performing the operation."
        );
        
        throw error; // Re-throw to allow caller to handle if needed
      }
    };
  }
  
  // Public API
  return {
    /**
     * Initiates the creation of the financial overview sheet.
     * Calls `FinancialPlanner.FinanceOverview.create()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.FinanceOverview.create()`.
     * @throws {Error} If an error occurs during overview generation.
     */
    createFinancialOverview: wrapWithFeedback(
      function() {
        return FinancialPlanner.FinanceOverview.create();
      },
      "Generating financial overview...",
      "Financial overview generated successfully!",
      "Failed to generate financial overview"
    ),
    
    /**
     * Generates the monthly spending report.
     * Calls `FinancialPlanner.MonthlySpendingReport.generate()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.MonthlySpendingReport.generate()`.
     * @throws {Error} If an error occurs during report generation.
     */
    generateMonthlySpendingReport: wrapWithFeedback(
      function() {
        return FinancialPlanner.MonthlySpendingReport.generate();
      },
      "Generating monthly spending report...",
      "Monthly spending report generated successfully!",
      "Failed to generate monthly spending report"
    ),
    
    /**
     * Displays key financial metrics.
     * Calls `FinancialPlanner.FinancialAnalysisService.showKeyMetrics()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.FinancialAnalysisService.showKeyMetrics()`.
     * @throws {Error} If an error occurs while displaying metrics.
     */
    showKeyMetrics: wrapWithFeedback(
      function() {
        return FinancialPlanner.FinancialAnalysisService.showKeyMetrics();
      },
      "Analyzing financial data...",
      "Key metrics displayed successfully!",
      "Failed to display key metrics"
    ),
    
    /**
     * Generates the yearly summary report.
     * Calls `FinancialPlanner.ReportService.generateYearlySummary()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.ReportService.generateYearlySummary()`.
     * @throws {Error} If an error occurs during report generation.
     */
    generateYearlySummary: wrapWithFeedback(
      function() {
        return FinancialPlanner.ReportService.generateYearlySummary();
      },
      "Generating yearly summary report...",
      "Yearly summary report generated successfully!",
      "Failed to generate yearly summary report"
    ),
    
    /**
     * Generates the category breakdown report.
     * Calls `FinancialPlanner.ReportService.generateCategoryBreakdown()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.ReportService.generateCategoryBreakdown()`.
     * @throws {Error} If an error occurs during report generation.
     */
    generateCategoryBreakdown: wrapWithFeedback(
      function() {
        return FinancialPlanner.ReportService.generateCategoryBreakdown();
      },
      "Generating category breakdown report...",
      "Category breakdown report generated successfully!",
      "Failed to generate category breakdown report"
    ),
    
    /**
     * Generates the savings analysis report.
     * Calls `FinancialPlanner.ReportService.generateSavingsAnalysis()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.ReportService.generateSavingsAnalysis()`.
     * @throws {Error} If an error occurs during report generation.
     */
    generateSavingsAnalysis: wrapWithFeedback(
      function() {
        return FinancialPlanner.ReportService.generateSavingsAnalysis();
      },
      "Generating savings analysis report...",
      "Savings analysis report generated successfully!",
      "Failed to generate savings analysis report"
    ),

    /**
     * Creates and displays a spending trends chart.
     * Calls `FinancialPlanner.VisualizationService.createSpendingTrendsChart()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.VisualizationService.createSpendingTrendsChart()`.
     * @throws {Error} If an error occurs during chart creation.
     */
    createSpendingTrendsChart: wrapWithFeedback(
      function() {
        return FinancialPlanner.VisualizationService.createSpendingTrendsChart();
      },
      "Creating spending trends chart...",
      "Spending trends chart created successfully!",
      "Failed to create spending trends chart"
    ),

    /**
     * Creates and displays a budget vs. actual spending chart.
     * Calls `FinancialPlanner.VisualizationService.createBudgetVsActualChart()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.VisualizationService.createBudgetVsActualChart()`.
     * @throws {Error} If an error occurs during chart creation.
     */
    createBudgetVsActualChart: wrapWithFeedback(
      function() {
        return FinancialPlanner.VisualizationService.createBudgetVsActualChart();
      },
      "Creating budget vs actual chart...",
      "Budget vs actual chart created successfully!",
      "Failed to create budget vs actual chart"
    ),

    /**
     * Creates and displays an income vs. expenses chart.
     * Calls `FinancialPlanner.VisualizationService.createIncomeVsExpensesChart()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.VisualizationService.createIncomeVsExpensesChart()`.
     * @throws {Error} If an error occurs during chart creation.
     */
    createIncomeVsExpensesChart: wrapWithFeedback(
      function() {
        return FinancialPlanner.VisualizationService.createIncomeVsExpensesChart();
      },
      "Creating income vs expenses chart...",
      "Income vs expenses chart created successfully!",
      "Failed to create income vs expenses chart"
    ),

    /**
     * Creates and displays a category pie chart.
     * Calls `FinancialPlanner.VisualizationService.createCategoryPieChart()`.
     * Provides UI feedback during the process.
     * @return {any} The result from `FinancialPlanner.VisualizationService.createCategoryPieChart()`.
     * @throws {Error} If an error occurs during chart creation.
     */
    createCategoryPieChart: wrapWithFeedback(
      function() {
        return FinancialPlanner.VisualizationService.createCategoryPieChart();
      },
      "Creating category pie chart...",
      "Category pie chart created successfully!",
      "Failed to create category pie chart"
    ),
    
    /**
     * Toggles the visibility of sub-categories in the financial overview sheet.
     * Uses `FinancialPlanner.SettingsService.toggleShowSubCategories()` to update the preference
     * and then updates the sheet display accordingly.
     * Provides UI feedback during the process.
     * @return {boolean} The new state of the 'show sub-categories' preference.
     * @throws {Error} If an error occurs while updating preferences or sheet display.
     */
    toggleShowSubCategories: wrapWithFeedback(
      function() {
        // Use the SettingsService to toggle the sub-categories preference
        const newValue = FinancialPlanner.SettingsService.toggleShowSubCategories();
        
        // Get the active spreadsheet and overview sheet
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const overviewSheet = ss.getSheetByName(FinancialPlanner.Config.getSheetNames().OVERVIEW);
        
        // If the overview sheet exists, update it based on the new preference
        if (overviewSheet) {
          if (newValue) {
            overviewSheet.showColumns(3, 1); // Show sub-categories column
          } else {
            overviewSheet.hideColumns(3, 1); // Hide sub-categories column
          }
        }
        
        return newValue;
      },
      "Updating display preferences...",
      "Display preferences updated successfully!",
      "Failed to update display preferences"
    ),
    
    /**
     * Refreshes all application caches.
     * This includes invalidating the general `CacheService` and the `DropdownService` cache.
     * Provides UI feedback during the process.
     * @return {boolean} True if the cache refresh process was initiated.
     * @throws {Error} If an error occurs during cache invalidation.
     */
    refreshCache: wrapWithFeedback(
      function() {
        // Invalidate general cache
        if (FinancialPlanner.CacheService && FinancialPlanner.CacheService.invalidateAll) {
          FinancialPlanner.CacheService.invalidateAll();
        }
        // Refresh dropdown specific cache
        if (FinancialPlanner.DropdownService && FinancialPlanner.DropdownService.refreshCache) {
          FinancialPlanner.DropdownService.refreshCache(); // This already has UI feedback
          return true; // Assuming DropdownService.refreshCache handles its own success/error messages
        }
        return true; // Fallback if DropdownService not available or doesn't handle feedback itself
      },
      "Refreshing all caches...", // General message
      "Caches refreshed successfully!", // General success, DropdownService might show specific one
      "Failed to refresh one or more caches"
    ),
    
    /**
     * Handles the `onOpen` event for the Google Sheet.
     * This function is automatically triggered by Google Apps Script when the spreadsheet is opened.
     * It creates the custom 'Financial Tools' menu in the Google Sheets UI.
     * Errors during menu creation are logged to the console and `ErrorService` but do not show a UI notification
     * to avoid disruption when the spreadsheet is opened.
     */
    onOpen: function() {
      try {
        // Create the menu
        const ui = SpreadsheetApp.getUi();
        
        // Create enhanced Financial Tools menu with submenus and icons
        ui.createMenu('📊 Financial Tools')
          .addItem('📈 Generate Overview', 'FinancialPlanner.Controllers.createFinancialOverview')
          .addSeparator()
          .addSubMenu(ui.createMenu('📋 Reports')
            .addItem('📝 Monthly Spending Report', 'FinancialPlanner.Controllers.generateMonthlySpendingReport')
            .addItem('📅 Yearly Summary (Coming Soon)', 'FinancialPlanner.Controllers.generateYearlySummary')
            .addItem('🔍 Category Breakdown (Coming Soon)', 'FinancialPlanner.Controllers.generateCategoryBreakdown')
            .addItem('💰 Savings Analysis (Coming Soon)', 'FinancialPlanner.Controllers.generateSavingsAnalysis'))
          .addSeparator()
          .addSubMenu(ui.createMenu('📊 Visualizations (Coming Soon)')
            .addItem('📉 Spending Trends Chart (Coming Soon)', 'FinancialPlanner.Controllers.createSpendingTrendsChart')
            .addItem('⚖️ Budget vs Actual (Coming Soon)', 'FinancialPlanner.Controllers.createBudgetVsActualChart')
            .addItem('💹 Income vs Expenses (Coming Soon)', 'FinancialPlanner.Controllers.createIncomeVsExpensesChart')
            .addItem('🍩 Category Pie Chart (Coming Soon)', 'FinancialPlanner.Controllers.createCategoryPieChart'))
          .addSeparator()
          .addSubMenu(ui.createMenu('🧮 Financial Analysis')
            .addItem('📊 Key Metrics', 'FinancialPlanner.Controllers.showKeyMetrics')
            .addItem('💡 Suggest Savings Opportunities (Coming Soon)', 'FinancialPlanner.Controllers.suggestSavingsOpportunities')
            .addItem('⚠️ Spending Anomaly Detection (Coming Soon)', 'FinancialPlanner.Controllers.detectSpendingAnomalies')
            .addItem('📌 Fixed vs Variable Expenses (Coming Soon)', 'FinancialPlanner.Controllers.analyzeFixedVsVariableExpenses')
            .addItem('🔮 Cash Flow Forecast (Coming Soon)', 'FinancialPlanner.Controllers.generateCashFlowForecast'))
          .addSeparator()
          .addSubMenu(ui.createMenu('⚙️ Settings')
            .addItem('🔄 Toggle Sub-Categories in Overview', 'FinancialPlanner.Controllers.toggleShowSubCategories')
            .addItem('🎯 Set Budget Targets (Coming Soon)', 'FinancialPlanner.Controllers.setBudgetTargets')
            .addItem('📧 Setup Email Reports (Coming Soon)', 'FinancialPlanner.Controllers.setupEmailReports')
            .addItem('🔄 Refresh Cache', 'FinancialPlanner.Controllers.refreshCache'))
          .addToUi();
      } catch (error) {
        // Log the error but don't show a UI notification
        // (this would be disruptive when opening the spreadsheet)
        errorService.log(error);
        console.error("Failed to create menu:", error);
      }
    },
    
    /**
     * Handles the `onEdit` event for the Google Sheet.
     * This function is automatically triggered by Google Apps Script when a user edits any cell in the spreadsheet.
     * It dispatches the edit event to the appropriate handler based on the sheet name
     * (e.g., `FinancialPlanner.FinanceOverview.handleEdit` or `FinancialPlanner.DropdownService.handleEdit`).
     * Errors during event handling are logged to the console and `ErrorService` but do not show a UI notification
     * to avoid disruption during editing.
     * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event object provided by Google Apps Script.
     *        See {@link https://developers.google.com/apps-script/guides/triggers/events#edit_3}
     */
    onEdit: function(e) {
      try {
        // Get the sheet that was edited
        const sheet = e.range.getSheet();
        const sheetName = sheet.getName();
        
        // Dispatch to the appropriate handler based on the sheet name
        if (sheetName === config.getSheetNames().OVERVIEW) {
          // Use the FinanceOverview module's handleEdit method
          if (FinancialPlanner.FinanceOverview && FinancialPlanner.FinanceOverview.handleEdit) {
            FinancialPlanner.FinanceOverview.handleEdit(e);
          }
        } else if (sheetName === config.getSheetNames().TRANSACTIONS) {
          // Use the DropdownService module's handleEdit method
          if (FinancialPlanner.DropdownService && FinancialPlanner.DropdownService.handleEdit) {
            FinancialPlanner.DropdownService.handleEdit(e);
          }
        }
        // Add more handlers for other sheets as they are refactored
        
      } catch (error) {
        // Log the error but don't show a UI notification
        // (this would be disruptive during editing)
        errorService.log(error);
        console.error("Error handling edit event:", error);
      }
    }
  };
})(FinancialPlanner.Config, FinancialPlanner.UIService, FinancialPlanner.ErrorService);

// For backward compatibility, create global references to the controller functions
// These can be removed once all code has been updated to use the namespace.

/**
 * Global `onOpen` trigger function for Google Apps Script.
 * Calls `FinancialPlanner.Controllers.onOpen()`.
 * This function is automatically called by Google Apps Script when the spreadsheet is opened.
 * @global
 */
function onOpen() {
  return FinancialPlanner.Controllers.onOpen();
}

/**
 * Global `onEdit` trigger function for Google Apps Script.
 * Calls `FinancialPlanner.Controllers.onEdit(e)`.
 * This function is automatically called by Google Apps Script when a cell is edited.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event object.
 * @global
 */
function onEdit(e) {
  return FinancialPlanner.Controllers.onEdit(e);
}

/**
 * Global `refreshCache` function, primarily for backward compatibility or direct invocation.
 * Calls `FinancialPlanner.Controllers.refreshCache()`.
 * @return {boolean | undefined} The result from `FinancialPlanner.Controllers.refreshCache()`.
 * @global
 */
function refreshCache() {
  return FinancialPlanner.Controllers.refreshCache();
}
