/**
 * Financial Planning Tools - Error Service Tests
 *
 * Exercises the real FinancialPlanner.ErrorService against the in-memory
 * Apps Script fakes in gas-fakes.js: errors are written to the "Error Log"
 * sheet and surfaced through UIService (ui.alert "Error: <message>").
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Errors = FinancialPlanner.ErrorService;
  let gas;

  function setup() {
    gas = FakeGas.install(global);
  }

  function errorLog() {
    return gas.spreadsheet.getSheetByName(FinancialPlanner.Config.getSheetNames().ERROR_LOG);
  }

  /** UIService.showErrorNotification shows ui.alert("<title>: <message>"). */
  function lastErrorNotification() {
    const alert = gas.ui.alerts[gas.ui.alerts.length - 1];
    if (!alert) return null;
    const text = alert.message;
    const i = text.indexOf(': ');
    return { title: text.slice(0, i), message: text.slice(i + 2) };
  }

  T.registerTest("ErrorService", "log should write a header and a row to the Error Log sheet", function() {
    setup();
    Errors.log(Errors.create("Sheet logging test", { severity: "high", code: 7 }));
    const rows = errorLog()._values();
    T.assertDeepEquals(['Timestamp', 'Error Type', 'Message', 'Details'], rows[0], "Header row should be written once.");
    T.assertEquals(2, rows.length, "One error row should follow the header.");
    T.assertEquals("FinancialPlannerError", rows[1][1], "Error type should be recorded.");
    T.assertEquals("Sheet logging test", rows[1][2], "Message should be recorded.");
    T.assertDeepEquals({ severity: "high", code: 7 }, JSON.parse(rows[1][3]), "Details should be stored as JSON.");
    T.assertEquals('#F9BDBD', errorLog().getRange(2, 1).getBackground(), "High severity rows should be highlighted red.");
  });

  T.registerTest("ErrorService", "log should not throw when the sheet is unavailable", function() {
    setup();
    global.SpreadsheetApp.getActiveSpreadsheet = function() { throw new Error("No spreadsheet"); };
    Errors.log(Errors.create("Still fine"));
    T.assertTrue(true, "log() swallowed the sheet failure.");
  });

  T.registerTest("ErrorService", "create should return a FinancialPlannerError instance", function() {
    setup();
    const message = "Test error message";
    const details = { code: 123, severity: "high" };
    const error = Errors.create(message, details);

    T.assertTrue(error instanceof Error, "Error should be an instance of Error.");
    // T.assertTrue(error instanceof FinancialPlannerError, "Error should be an instance of FinancialPlannerError."); // Cannot test custom class instance directly this way in GAS
    T.assertEquals("FinancialPlannerError", error.name, "Error name should be 'FinancialPlannerError'.");
    T.assertEquals(message, error.message, "Error message should match.");
    T.assertDeepEquals(details, error.details, "Error details should match.");
    T.assertTrue(error.timestamp instanceof Date, "Error should have a timestamp.");
  });

  T.registerTest("ErrorService", "log should execute without throwing (verification via logs)", function() {
    setup();
    const error = Errors.create("Logging test", { data: "sample" });
    try {
      Errors.log(error);
      // If we reach here, log executed without throwing internal errors.
      T.assertTrue(true, "log() executed without throwing an error.");
    } catch (e) {
      T.assertTrue(false, `log() should not have thrown an error, but threw: ${e.message}`);
    }
  });

  T.registerTest("ErrorService", "handle should call log and uiService.showErrorNotification", function() {
    setup();
    const error = Errors.create("Handling test", { id: 456 });
    const userMessage = "Something went wrong during handling test.";

    try {
      Errors.handle(error, userMessage);
      // Check if UI service was called
      T.assertNotNull(lastErrorNotification(), "uiService.showErrorNotification should have been called.");
      T.assertEquals("Error", lastErrorNotification().title, "Notification title should be 'Error'.");
      T.assertEquals(userMessage, lastErrorNotification().message, "Notification message should match userFriendlyMessage.");
    } catch (e) {
       T.assertTrue(false, `handle() should not have thrown an error, but threw: ${e.message}`);
    }
  });
  
   T.registerTest("ErrorService", "handle should use error message if user message not provided", function() {
    setup();
    const error = Errors.create("Error message only");

    try {
      Errors.handle(error); // No userFriendlyMessage
      // Check if UI service was called
      T.assertNotNull(lastErrorNotification(), "uiService.showErrorNotification should have been called.");
      T.assertEquals("Error", lastErrorNotification().title, "Notification title should be 'Error'.");
      T.assertEquals(error.message, lastErrorNotification().message, "Notification message should default to error.message.");
    } catch (e) {
       T.assertTrue(false, `handle() should not have thrown an error, but threw: ${e.message}`);
    }
  });

  T.registerTest("ErrorService", "wrap should execute function normally if no error", function() {
    setup();
    let executed = false;
    const wrappedFunc = Errors.wrap(function(a, b) {
      executed = true;
      return a + b;
    }, "Wrapper test failed");

    const result = wrappedFunc(5, 3);
    T.assertTrue(executed, "Wrapped function should have been executed.");
    T.assertEquals(8, result, "Wrapped function should return the correct result.");
    T.assertTrue(lastErrorNotification() === null, "showErrorNotification should not have been called.");
  });

  T.registerTest("ErrorService", "wrap should handle error and re-throw", function() {
    setup();
    const errorMessage = "Intentional error";
    const userMessage = "Wrapped function failed as expected.";
    const wrappedFunc = Errors.wrap(function() {
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
    T.assertNotNull(lastErrorNotification(), "showErrorNotification should have been called.");
    T.assertEquals(userMessage, lastErrorNotification().message, "Notification should use the user-friendly message.");
  });
  
   T.registerTest("ErrorService", "wrap should use default user message if none provided", function() {
    setup();
    const errorMessage = "Another intentional error";
    const wrappedFunc = Errors.wrap(function() {
      throw new Error(errorMessage);
    }); // No userFriendlyMessage

    let caughtError = null;
    try {
      wrappedFunc();
    } catch (e) {
      caughtError = e;
    }

    T.assertNotNull(caughtError, "wrap should have re-thrown the error.");
    T.assertNotNull(lastErrorNotification(), "showErrorNotification should have been called.");
    T.assertEquals("An error occurred while performing the operation.", lastErrorNotification().message, "Notification should use the default user-friendly message.");
  });


})();
