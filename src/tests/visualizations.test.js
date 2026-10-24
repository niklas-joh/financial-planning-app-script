/**
 * Financial Planning Tools - Visualization Service Tests
 *
 * Exercises the real FinancialPlanner.VisualizationService against the
 * in-memory Apps Script fakes in gas-fakes.js.
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Viz = FinancialPlanner.VisualizationService;
  let gas;

  [
    ['createSpendingTrendsChart', 'Spending Trends Chart - Coming Soon!'],
    ['createBudgetVsActualChart', 'Budget vs Actual Chart - Coming Soon!'],
    ['createIncomeVsExpensesChart', 'Income vs Expenses Chart - Coming Soon!'],
    ['createCategoryPieChart', 'Category Pie Chart - Coming Soon!'],
  ].forEach(function([method, message]) {
    T.registerTest("VisualizationService", method + " should show 'Coming Soon' alert", function() {
      gas = FakeGas.install(global);
      Viz[method]();

      T.assertEquals(1, gas.ui.alerts.length, 'Exactly one alert should be shown.');
      T.assertEquals(message, gas.ui.alerts[0].message, 'Alert message should be correct.');
      const toasts = gas.spreadsheet.toasts;
      T.assertEquals('', toasts[toasts.length - 1].message, 'The loading toast should be cleared.');
    });

    T.registerTest("VisualizationService", "global " + method + " should delegate to the service", function() {
      gas = FakeGas.install(global);
      global[method]();
      T.assertEquals(message, gas.ui.alerts[0].message, 'Global wrapper should show the same alert.');
    });
  });
})();
