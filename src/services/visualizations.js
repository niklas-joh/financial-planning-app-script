/**
 * Financial Planning Tools - Visualization Service
 *
 * This file provides chart generation functionality for the Financial Planning Tools project.
 * It follows the namespace pattern and uses dependency injection for better maintainability.
 */

// Create the VisualizationService module within the FinancialPlanner namespace
FinancialPlanner.VisualizationService = (function(utils, uiService, errorService, config) {
  // Private variables and functions (if any in the future)

  // Public API
  return {
    /**
     * Creates a spending trends chart
     */
    createSpendingTrendsChart: function() {
      try {
        uiService.showLoadingSpinner("Creating spending trends chart...");
        // TODO: Implement spending trends chart
        SpreadsheetApp.getUi().alert('Spending Trends Chart - Coming Soon!');
        uiService.hideLoadingSpinner();
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Failed to create spending trends chart");
      }
    },

    /**
     * Creates a budget vs actual chart
     */
    createBudgetVsActualChart: function() {
      try {
        uiService.showLoadingSpinner("Creating budget vs actual chart...");
        // TODO: Implement budget vs actual chart
        SpreadsheetApp.getUi().alert('Budget vs Actual Chart - Coming Soon!');
        uiService.hideLoadingSpinner();
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Failed to create budget vs actual chart");
      }
    },

    /**
     * Creates an income vs expenses chart
     */
    createIncomeVsExpensesChart: function() {
      try {
        uiService.showLoadingSpinner("Creating income vs expenses chart...");
        // TODO: Implement income vs expenses chart
        SpreadsheetApp.getUi().alert('Income vs Expenses Chart - Coming Soon!');
        uiService.hideLoadingSpinner();
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Failed to create income vs expenses chart");
      }
    },

    /**
     * Creates a category pie chart
     */
    createCategoryPieChart: function() {
      try {
        uiService.showLoadingSpinner("Creating category pie chart...");
        // TODO: Implement category pie chart
        SpreadsheetApp.getUi().alert('Category Pie Chart - Coming Soon!');
        uiService.hideLoadingSpinner();
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Failed to create category pie chart");
      }
    }
  };
})(FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.ErrorService, FinancialPlanner.Config);

// Backward compatibility layer for existing global functions
function createSpendingTrendsChart() {
  return FinancialPlanner.VisualizationService.createSpendingTrendsChart();
}

function createBudgetVsActualChart() {
  return FinancialPlanner.VisualizationService.createBudgetVsActualChart();
}

function createIncomeVsExpensesChart() {
  return FinancialPlanner.VisualizationService.createIncomeVsExpensesChart();
}

function createCategoryPieChart() {
  return FinancialPlanner.VisualizationService.createCategoryPieChart();
}
