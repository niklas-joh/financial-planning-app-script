/**
 * Financial Planning Tools - Report Service Tests
 *
 * Exercises the real FinancialPlanner.ReportService against the in-memory
 * Apps Script fakes in gas-fakes.js.
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Reports = FinancialPlanner.ReportService;
  let gas;

  function setup() {
    gas = FakeGas.install(global);
  }

  [
    ['generateYearlySummary', 'Yearly Summary', 'Generating yearly summary report...'],
    ['generateCategoryBreakdown', 'Category Breakdown', 'Generating category breakdown report...'],
    ['generateSavingsAnalysis', 'Savings Analysis', 'Generating savings analysis report...'],
  ].forEach(function([method, title, spinner]) {
    T.registerTest("ReportService", method + " should show 'Coming Soon' alert", function() {
      setup();
      const result = Reports[method]();

      T.assertTrue(result === undefined, 'Placeholder should return undefined.');
      T.assertEquals(1, gas.ui.alerts.length, 'Exactly one alert should be shown.');
      T.assertEquals(title, gas.ui.alerts[0].title, 'Alert title should match.');
      T.assertEquals('Coming Soon!', gas.ui.alerts[0].message, 'Alert message should match.');
      T.assertEquals(gas.ui.ButtonSet.OK, gas.ui.alerts[0].buttons, 'Alert should use the OK button set.');
    });

    T.registerTest("ReportService", method + " should show and then clear the loading toast", function() {
      setup();
      Reports[method]();

      const toasts = gas.spreadsheet.toasts;
      T.assertEquals(2, toasts.length, 'A spinner toast and a clearing toast are expected.');
      T.assertEquals(spinner, toasts[0].message, 'Spinner message should match.');
      T.assertEquals('', toasts[1].message, 'The second toast clears the spinner.');
    });
  });

  T.registerTest("ReportService", "errors are routed to ErrorService and the spinner is cleared", function() {
    setup();
    // Make the alert fail; the service must catch it, report it and return null.
    gas.ui.alert = function() { throw new Error('UI unavailable'); };
    const original = FinancialPlanner.ErrorService.handle;
    let handled = null;
    FinancialPlanner.ErrorService.handle = function(error, message) { handled = { error, message }; };
    try {
      const result = Reports.generateYearlySummary();
      T.assertEquals(null, result, 'Should return null on error.');
      T.assertNotNull(handled, 'ErrorService.handle should be called.');
      T.assertEquals('Failed to generate yearly summary report', handled.message, 'User-facing message should match.');
      T.assertEquals('', gas.spreadsheet.toasts[gas.spreadsheet.toasts.length - 1].message, 'Spinner should be cleared.');
    } finally {
      FinancialPlanner.ErrorService.handle = original;
    }
  });
})();
