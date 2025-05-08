/**
 * Financial Planning Tools - Error Service Tests
 *
 * This file contains tests for the FinancialPlanner.ErrorService module.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;
  const ErrorService = FinancialPlanner.ErrorService;

  // --- Mock Dependencies ---
  // Simple mock for Config to provide sheet names
  const mockConfig = {
    getSheetNames: function() {
      return { ERROR_LOG: "Mock Error Log" };
    },
    getSection: function(section) { // Added for compatibility if ErrorService uses it
        if (section === 'SHEETS') return this.getSheetNames();
        return {};
    }
  };

  // Simple mock for UIService to track notifications
  let lastErrorNotification = null;
  const mockUiService = {
    showErrorNotification: function(title, message) {
      lastErrorNotification = { title: title, message: message };
      // console.log(`Mock UI Service: showErrorNotification called with title "${title}", message "${message}"`);
    },
    // Add other methods if needed by ErrorService, though unlikely for basic tests
    showLoadingSpinner: function() {},
    hideLoadingSpinner: function() {},
    showSuccessNotification: function() {}
  };
  
  // Simple mock for Utils used by logToSheet
  const mockUtils = {
      getOrCreateSheet: function(ss, sheetName) {
          // Return a mock sheet object that supports basic logging needs
          return {
              appendRow: function(rowData) { /* console.log(`Mock Sheet (${sheetName}): Appending row: ${JSON.stringify(rowData)}`); */ },
              getRange: function() { 
                  return { 
                      setFontWeight: function() { return this; }, 
                      setNumberFormat: function() { return this; },
                      setBackground: function() { return this; }
                  }; 
              },
              getLastRow: function() { return 1; } // Simulate existing header
          };
      }
  };

  // --- Test Suite Setup ---
  // Inject mocks - This assumes ErrorService is defined using the IIFE pattern
  // We need to redefine it here with mocks. This is a limitation of simple GAS testing.
  // Ideally, a dependency injection framework or more advanced mocking would be used.
  // For now, we redefine the service for testing purposes.
  const TestErrorService = (function(config, uiService, utils) {
     // --- Copy of ErrorService Implementation Start ---
     // (Paste the implementation of FinancialPlanner.ErrorService here,
     // ensuring it uses the passed-in config, uiService, and utils mocks)
      class FinancialPlannerError extends Error {
        constructor(message, details = {}) {
          super(message);
          this.name = 'FinancialPlannerError';
          this.details = details;
          this.timestamp = new Date();
        }
      }
      
      function logToSheet(error) {
        try {
          const ss = SpreadsheetApp.getActiveSpreadsheet(); // Still uses global SS object
          const errorSheet = utils.getOrCreateSheet(ss, config.getSheetNames().ERROR_LOG); // Use mocked utils
          
          if (errorSheet.getLastRow() === 0) {
            errorSheet.appendRow(["Timestamp", "Error Type", "Message", "Details"]);
            errorSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
          }
          
          const errorDetails = error.details || {};
          const formattedDetails = JSON.stringify(errorDetails);
          
          errorSheet.appendRow([
            error.timestamp || new Date(), 
            error.name || "Error", 
            error.message, 
            formattedDetails
          ]);
          
          const lastRow = errorSheet.getLastRow();
          errorSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
          
          const severity = errorDetails.severity || "low";
          const bgColor = severity === "high" ? "#F9BDBD" : 
                          severity === "medium" ? "#FFE0B2" : "#E1F5FE";
          errorSheet.getRange(lastRow, 1, 1, 4).setBackground(bgColor);
        } catch (logError) {
          console.error("Failed to log error to sheet:", logError);
          console.error("Original error:", error.message, error.details);
        }
      }
      
      function logToConsole(error) {
        // In tests, Logger.log is often preferred over console.error for GAS execution logs
        Logger.log(`[${error.name || "Error"}] ${error.message}`);
        if (error.details) {
          Logger.log("Details: " + JSON.stringify(error.details));
        }
        if (error.stack) {
           Logger.log("Stack trace: " + error.stack);
        }
      }
      
      return {
        create: function(message, details = {}) {
          return new FinancialPlannerError(message, details);
        },
        log: function(error) {
          logToConsole(error);
          logToSheet(error); // Will use mocked utils/config
        },
        handle: function(error, userFriendlyMessage) {
          this.log(error);
          uiService.showErrorNotification( // Use mocked uiService
            "Error",
            userFriendlyMessage || error.message
          );
        },
        wrap: function(fn, userFriendlyMessage) {
          const self = this; 
          return function() {
            try {
              return fn.apply(this, arguments);
            } catch (error) {
              self.handle( 
                error,
                userFriendlyMessage || "An error occurred while performing the operation."
              );
              throw error; 
            }
          };
        }
      };
     // --- Copy of ErrorService Implementation End ---
  })(mockConfig, mockUiService, mockUtils); // Pass mocks

  // --- Test Cases ---

  T.registerTest("ErrorService", "create should return a FinancialPlannerError instance", function() {
    const message = "Test error message";
    const details = { code: 123, severity: "high" };
    const error = TestErrorService.create(message, details);

    T.assertTrue(error instanceof Error, "Error should be an instance of Error.");
    // T.assertTrue(error instanceof FinancialPlannerError, "Error should be an instance of FinancialPlannerError."); // Cannot test custom class instance directly this way in GAS
    T.assertEquals("FinancialPlannerError", error.name, "Error name should be 'FinancialPlannerError'.");
    T.assertEquals(message, error.message, "Error message should match.");
    T.assertDeepEquals(details, error.details, "Error details should match.");
    T.assertTrue(error.timestamp instanceof Date, "Error should have a timestamp.");
  });

  T.registerTest("ErrorService", "log should execute without throwing (verification via logs)", function() {
    // We can't directly verify sheet/console output easily in automated tests.
    // We just check that calling log doesn't throw an unexpected error itself.
    const error = TestErrorService.create("Logging test", { data: "sample" });
    try {
      TestErrorService.log(error);
      // If we reach here, log executed without throwing internal errors.
      T.assertTrue(true, "log() executed without throwing an error.");
    } catch (e) {
      T.assertTrue(false, `log() should not have thrown an error, but threw: ${e.message}`);
    }
  });

  T.registerTest("ErrorService", "handle should call log and uiService.showErrorNotification", function() {
    lastErrorNotification = null; // Reset mock tracker
    const error = TestErrorService.create("Handling test", { id: 456 });
    const userMessage = "Something went wrong during handling test.";

    try {
      TestErrorService.handle(error, userMessage);
      // Check if UI service was called
      T.assertNotNull(lastErrorNotification, "uiService.showErrorNotification should have been called.");
      T.assertEquals("Error", lastErrorNotification.title, "Notification title should be 'Error'.");
      T.assertEquals(userMessage, lastErrorNotification.message, "Notification message should match userFriendlyMessage.");
    } catch (e) {
       T.assertTrue(false, `handle() should not have thrown an error, but threw: ${e.message}`);
    }
  });
  
   T.registerTest("ErrorService", "handle should use error message if user message not provided", function() {
    lastErrorNotification = null; // Reset mock tracker
    const error = TestErrorService.create("Error message only");

    try {
      TestErrorService.handle(error); // No userFriendlyMessage
      // Check if UI service was called
      T.assertNotNull(lastErrorNotification, "uiService.showErrorNotification should have been called.");
      T.assertEquals("Error", lastErrorNotification.title, "Notification title should be 'Error'.");
      T.assertEquals(error.message, lastErrorNotification.message, "Notification message should default to error.message.");
    } catch (e) {
       T.assertTrue(false, `handle() should not have thrown an error, but threw: ${e.message}`);
    }
  });

  T.registerTest("ErrorService", "wrap should execute function normally if no error", function() {
    lastErrorNotification = null;
    let executed = false;
    const wrappedFunc = TestErrorService.wrap(function(a, b) {
      executed = true;
      return a + b;
    }, "Wrapper test failed");

    const result = wrappedFunc(5, 3);
    T.assertTrue(executed, "Wrapped function should have been executed.");
    T.assertEquals(8, result, "Wrapped function should return the correct result.");
    T.assertTrue(lastErrorNotification === null, "showErrorNotification should not have been called.");
  });

  T.registerTest("ErrorService", "wrap should handle error and re-throw", function() {
    lastErrorNotification = null;
    const errorMessage = "Intentional error";
    const userMessage = "Wrapped function failed as expected.";
    const wrappedFunc = TestErrorService.wrap(function() {
      throw new Error(errorMessage);
    }, userMessage);

    let caughtError = null;
    try {
      wrappedFunc();
    } catch (e) {
      caughtError = e;
    }

    T.assertNotNull(caughtError, "wrap should have re-thrown the error.");
    T.assertEquals(errorMessage, caughtError.message, "The original error message should be preserved.");
    T.assertNotNull(lastErrorNotification, "showErrorNotification should have been called.");
    T.assertEquals(userMessage, lastErrorNotification.message, "Notification should use the user-friendly message.");
  });
  
   T.registerTest("ErrorService", "wrap should use default user message if none provided", function() {
    lastErrorNotification = null;
    const errorMessage = "Another intentional error";
    const wrappedFunc = TestErrorService.wrap(function() {
      throw new Error(errorMessage);
    }); // No userFriendlyMessage

    let caughtError = null;
    try {
      wrappedFunc();
    } catch (e) {
      caughtError = e;
    }

    T.assertNotNull(caughtError, "wrap should have re-thrown the error.");
    T.assertNotNull(lastErrorNotification, "showErrorNotification should have been called.");
    T.assertEquals("An error occurred while performing the operation.", lastErrorNotification.message, "Notification should use the default user-friendly message.");
  });


})(); // End IIFE
