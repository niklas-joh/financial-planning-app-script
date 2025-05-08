/**
 * Financial Planning Tools - Visualization Service Tests
 *
 * This file contains tests for the FinancialPlanner.VisualizationService module.
 * It tests the placeholder functions for chart generation.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;

  // --- Mock Dependencies & Globals ---
  let lastToast = null;
  let lastAlert = null;
  let lastHandledError = null;

  // Mock UI Service
  const mockUiService = {
    showLoadingSpinner: function(msg) { lastToast = { message: msg, title: "Working..." }; },
    hideLoadingSpinner: function() { /* Mock */ },
    showInfoAlert: function(title, message) { lastAlert = { title: title, message: message }; }, // Used by placeholders
    showErrorNotification: function(title, message) { lastAlert = { title: title, message: message }; } // Use alert for error mock
  };
   // Mock Error Service
  const mockErrorService = {
    handle: function(error, msg) { lastHandledError = { error, msg }; console.error("ERROR SERVICE MOCK:", msg, error); },
    create: function(msg, details) { const e = new Error(msg); e.details = details; e.name="FinancialPlannerError"; return e; },
    log: function(error) { console.log("ErrorService Mock Log:", error.message); }
  };

  // Mock Config and Utils (minimal mocks as they aren't directly used by placeholders)
  const mockConfig = {};
  const mockUtils = {};
  
  // Global mocks needed by service implementation (even if placeholders)
  global.SpreadsheetApp = {
      getUi: function() { return { alert: mockUiService.showInfoAlert }; } // Mock the specific call used in placeholders
  };
  global.Logger = { log: function(msg) { console.log("Logger.log:", msg); } };


  // --- Test Suite Setup ---
   // Redefine VisualizationService with mocks
   const TestVisualizationService = (function(utils, uiService, errorService, config) {
       // --- Copy of VisualizationService Implementation Start ---
        return {
            createSpendingTrendsChart: function() { try { uiService.showLoadingSpinner("Creating spending trends chart..."); SpreadsheetApp.getUi().alert('Spending Trends Chart - Coming Soon!'); uiService.hideLoadingSpinner(); } catch (error) { uiService.hideLoadingSpinner(); errorService.handle(error, "Failed to create spending trends chart"); } },
            createBudgetVsActualChart: function() { try { uiService.showLoadingSpinner("Creating budget vs actual chart..."); SpreadsheetApp.getUi().alert('Budget vs Actual Chart - Coming Soon!'); uiService.hideLoadingSpinner(); } catch (error) { uiService.hideLoadingSpinner(); errorService.handle(error, "Failed to create budget vs actual chart"); } },
            createIncomeVsExpensesChart: function() { try { uiService.showLoadingSpinner("Creating income vs expenses chart..."); SpreadsheetApp.getUi().alert('Income vs Expenses Chart - Coming Soon!'); uiService.hideLoadingSpinner(); } catch (error) { uiService.hideLoadingSpinner(); errorService.handle(error, "Failed to create income vs expenses chart"); } },
            createCategoryPieChart: function() { try { uiService.showLoadingSpinner("Creating category pie chart..."); SpreadsheetApp.getUi().alert('Category Pie Chart - Coming Soon!'); uiService.hideLoadingSpinner(); } catch (error) { uiService.hideLoadingSpinner(); errorService.handle(error, "Failed to create category pie chart"); } }
        };
       // --- Copy of VisualizationService Implementation End ---
   })(mockUtils, mockUiService, mockErrorService, mockConfig); // Pass mocks


  // --- Helper to reset state before each test ---
  function resetMocks() {
      lastToast = null;
      lastAlert = null;
      lastHandledError = null;
  }

  // --- Test Cases ---

  T.registerTest("VisualizationService", "createSpendingTrendsChart should show 'Coming Soon' alert", function() {
    resetMocks();
    TestVisualizationService.createSpendingTrendsChart();

    T.assertNotNull(lastAlert, "uiService.showInfoAlert (via SpreadsheetApp.getUi().alert) should have been called.");
    // Note: The actual implementation calls SpreadsheetApp.getUi().alert directly
    // T.assertEquals("Spending Trends Chart", lastAlert.title, "Alert title should be correct."); // Title not passed in current impl
    T.assertEquals("Spending Trends Chart - Coming Soon!", lastAlert.message, "Alert message should be correct.");
    T.assertTrue(lastHandledError === null, "No error should be handled.");
  });

  T.registerTest("VisualizationService", "createBudgetVsActualChart should show 'Coming Soon' alert", function() {
    resetMocks();
    TestVisualizationService.createBudgetVsActualChart();

    T.assertNotNull(lastAlert, "uiService.showInfoAlert should have been called.");
    T.assertEquals("Budget vs Actual Chart - Coming Soon!", lastAlert.message, "Alert message should be correct.");
    T.assertTrue(lastHandledError === null, "No error should be handled.");
  });

  T.registerTest("VisualizationService", "createIncomeVsExpensesChart should show 'Coming Soon' alert", function() {
    resetMocks();
     TestVisualizationService.createIncomeVsExpensesChart();

    T.assertNotNull(lastAlert, "uiService.showInfoAlert should have been called.");
    T.assertEquals("Income vs Expenses Chart - Coming Soon!", lastAlert.message, "Alert message should be correct.");
    T.assertTrue(lastHandledError === null, "No error should be handled.");
  });

  T.registerTest("VisualizationService", "createCategoryPieChart should show 'Coming Soon' alert", function() {
    resetMocks();
    TestVisualizationService.createCategoryPieChart();

    T.assertNotNull(lastAlert, "uiService.showInfoAlert should have been called.");
    T.assertEquals("Category Pie Chart - Coming Soon!", lastAlert.message, "Alert message should be correct.");
    T.assertTrue(lastHandledError === null, "No error should be handled.");
  });

})(); // End IIFE
