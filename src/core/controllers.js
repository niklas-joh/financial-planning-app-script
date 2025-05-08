/**
 * Financial Planning Tools - Controllers
 * 
 * This file provides a centralized set of controller functions that serve as
 * entry points for UI-triggered actions. These functions coordinate between
 * the UI and the underlying services.
 */

// Create the Controllers module within the FinancialPlanner namespace
FinancialPlanner.Controllers = (function(config, uiService, errorService) {
  // Private variables and functions
  
  /**
   * Wraps a controller function with standard error handling and UI feedback
   * @param {Function} fn - The function to wrap
   * @param {String} startMessage - Message to show when the operation starts
   * @param {String} successMessage - Message to show when the operation succeeds
   * @param {String} errorMessage - Message to show when the operation fails
   * @return {Function} The wrapped function
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
     * Creates the financial overview
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
     * Generates a monthly spending report
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
     * Shows key financial metrics
     */
    showKeyMetrics: wrapWithFeedback(
      function() {
        return FinancialPlanner.FinancialAnalysis.showKeyMetrics();
      },
      "Analyzing financial data...",
      "Key metrics displayed successfully!",
      "Failed to display key metrics"
    ),
    
    /**
     * Generates a yearly summary report
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
     * Generates a category breakdown report
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
     * Generates a savings analysis report
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
     * Creates a spending trends chart
     */
    createSpendingTrendsChart: wrapWithFeedback(
      function() {
        return FinancialPlanner.VisualizationService.createSpendingTrendsChart();
      },
      "Creating spending trends chart...",
      "Spending trends chart created successfully!", // Or appropriate message
      "Failed to create spending trends chart"
    ),

    /**
     * Creates a budget vs actual chart
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
     * Creates an income vs expenses chart
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
     * Creates a category pie chart
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
     * Toggles the display of sub-categories in the overview
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
     * Refreshes the cache
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
     * Handles the onOpen event
     * This function is called when the spreadsheet is opened
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
     * Handles the onEdit event
     * This function is called when the user edits the spreadsheet
     * @param {Object} e - The edit event object
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
// These can be removed once all code has been updated to use the namespace
function onOpen() {
  return FinancialPlanner.Controllers.onOpen();
}

function onEdit(e) {
  return FinancialPlanner.Controllers.onEdit(e);
}

function refreshCache() {
  return FinancialPlanner.Controllers.refreshCache();
}
