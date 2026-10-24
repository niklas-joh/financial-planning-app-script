/**
 * Financial Planning Tools - Monthly Spending Report Tests
 *
 * Exercises the real FinancialPlanner.MonthlySpendingReport against the
 * in-memory Apps Script fakes. Transactions are generated relative to today,
 * because the report always covers the current month.
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Report = FinancialPlanner.MonthlySpendingReport;
  let gas;

  function monthsAgo(n, day) {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth() - n, day || 10);
  }

  function setup(rows) {
    gas = FakeGas.install(global);
    const tx = gas.spreadsheet.insertSheet('Transactions');
    const data = [['Date', 'Description', 'Type', 'Category', 'Sub-Category', 'Amount']].concat(rows || [
      [monthsAgo(0), 'Supermarket', 'Essentials', 'Food', 'Groceries', -100],
      [monthsAgo(0), 'Pizza', 'Wants/Pleasure', 'Food', 'Restaurants', -50],
      [monthsAgo(0), 'Landlord', 'Essentials', 'Housing', 'Rent', -850],
      [monthsAgo(0), 'Employer', 'Income', 'Salary', '', 3000],
      [monthsAgo(1), 'Supermarket', 'Essentials', 'Food', 'Groceries', -50],
      [monthsAgo(1), 'Landlord', 'Essentials', 'Housing', 'Rent', -850],
      [monthsAgo(2), 'Landlord', 'Essentials', 'Housing', 'Rent', -830],
      [monthsAgo(5), 'Supermarket', 'Essentials', 'Food', 'Groceries', -999], // outside the 3-month window
    ]);
    tx.getRange(1, 1, data.length, 6).setValues(data);
  }

  /** Report rows keyed by "Category/Sub-Category" (sub-rows) or category name (header rows). */
  function reportRows(sheet) {
    const rows = {};
    let category = null;
    let done = false;
    sheet._values().slice(3).forEach(function(r) {
      // The chart's data block follows the TOTAL row; stop there.
      if (done) return;
      if (r[0] === 'TOTAL EXPENSES') { rows[r[0]] = r; done = true; return; }
      if (r[0]) { category = r[0]; rows[category] = r; }
      else if (r[1] && category) rows[category + '/' + r[1]] = r;
    });
    return rows;
  }

  T.registerTest("MonthlySpendingReport", "generate should create report sheet with title", function() {
    setup();
    const sheet = Report.generate();
    T.assertNotNull(sheet, "generate() should return a sheet object.");
    T.assertEquals('Monthly Report', sheet.getName(), "Sheet should be named 'Monthly Report'.");
    const now = new Date();
    T.assertEquals('Monthly Spending Report - ' + FinancialPlanner.Utils.getMonthName(now.getMonth()) + ' ' + now.getFullYear(),
      sheet.getRange('A1').getValue(), "Title should name the current month.");
    T.assertDeepEquals(['Category', 'Sub-Category', 'Amount', '% of Total', 'Avg Last 3 Months', 'Trend'],
      sheet.getRange('A3:F3').getValues()[0], "Header row should be written.");
    T.assertEquals('Monthly spending report generated!', gas.spreadsheet.toasts.pop().message, "Success toast should be shown.");
  });

  T.registerTest("MonthlySpendingReport", "generate should aggregate expenses correctly", function() {
    setup();
    const rows = reportRows(Report.generate());
    T.assertNotNull(rows['Food/Groceries'], "Food/Groceries row should exist.");
    T.assertEquals(150, rows['Food'][2], "Food total should include expenses and wants.");
    T.assertEquals(100, rows['Food/Groceries'][2], "Amounts should be absolute values.");
    T.assertEquals(850, rows['Housing/Rent'][2], "Rent should be this month's amount only.");
    T.assertTrue(rows['Salary'] === undefined, "Income should be excluded.");
    T.assertEquals(1000, rows['TOTAL EXPENSES'][2], "Total should sum all expense rows.");
  });

  T.registerTest("MonthlySpendingReport", "generate should count every configured expense type", function() {
    setup([
      [monthsAgo(0), 'Rent', 'Essentials', 'Housing', 'Rent', -800],
      [monthsAgo(0), 'Cinema', 'Wants/Pleasure', 'Fun', 'Cinema', -20],
      [monthsAgo(0), 'Gift', 'Extra', 'Gifts', '', -30],
      [monthsAgo(0), 'ETF', 'Savings', 'Investments', '', -500],
    ]);
    const rows = reportRows(Report.generate());
    T.assertEquals(850, rows['TOTAL EXPENSES'][2], "Essentials, Wants/Pleasure and Extra count; Savings does not.");
    T.assertNotNull(rows['Gifts/(None)'], "Missing sub-categories are shown as (None).");
  });

  T.registerTest("MonthlySpendingReport", "generate should calculate percentages correctly", function() {
    setup();
    const rows = reportRows(Report.generate());
    T.assertEquals(0.15, rows['Food'][3], "Food share should be 15%.");
    T.assertEquals(0.85, rows['Housing/Rent'][3], "Rent share should be 85%.");
    T.assertEquals(1, rows['TOTAL EXPENSES'][3], "Total share should be 100%.");
  });

  T.registerTest("MonthlySpendingReport", "generate should compute 3-month averages and trends", function() {
    setup();
    const rows = reportRows(Report.generate());
    T.assertEquals(50, rows['Food/Groceries'][4], "Average uses only months with data in the last 3.");
    T.assertEquals('↑ 100%', rows['Food/Groceries'][5], "Doubling should show an upward trend.");
    T.assertEquals(840, rows['Housing/Rent'][4], "Rent average over two months.");
    T.assertEquals('→ Stable', rows['Housing/Rent'][5], "Within 10% should be stable.");
    T.assertEquals('', rows['Food/Restaurants'][5], "No history means no trend.");
  });

  T.registerTest("MonthlySpendingReport", "generate should handle missing Transactions sheet", function() {
    gas = FakeGas.install(global);
    const result = Report.generate();
    T.assertEquals(null, result, "generate() should return null on error.");
    const alert = gas.ui.alerts[gas.ui.alerts.length - 1];
    T.assertTrue(alert.message.includes('Failed to generate monthly spending report'), "User should see the failure message.");
  });

  T.registerTest("MonthlySpendingReport", "generate should reject a sheet without the required columns", function() {
    gas = FakeGas.install(global);
    gas.spreadsheet.insertSheet('Transactions').getRange(1, 1, 1, 2).setValues([['When', 'What']]);
    T.assertEquals(null, Report.generate(), "generate() should return null.");
    const log = gas.spreadsheet.getSheetByName('Error Log')._values();
    T.assertEquals('Could not find required columns in Transaction sheet', log[log.length - 1][2], "Cause should be logged.");
  });

  T.registerTest("MonthlySpendingReport", "amount columns use the configured currency format", function() {
    setup();
    const sheet = Report.generate();
    const expected = FinancialPlanner.Config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT;
    const formats = sheet._calls.filter((c) => c.method === 'setNumberFormat').map((c) => c.args[0]);
    T.assertEquals(2, formats.filter((f) => f === expected).length, "Amount and average columns should get the currency format.");
    T.assertTrue(formats.indexOf('€') === -1, "The bare currency symbol must not be used as a number format.");
  });

  T.registerTest("MonthlySpendingReport", "generate should add charts", function() {
    setup();
    const sheet = Report.generate();
    const charts = sheet.getCharts();
    T.assertEquals(1, charts.length, "One chart should have been added.");
    T.assertEquals('PIE', charts[0].spec.type, "It should be a pie chart.");
    T.assertEquals('Expense Breakdown by Category', charts[0].spec.options.title, "Chart title should match.");
  });

  T.registerTest("MonthlySpendingReport", "regenerating replaces the previous report", function() {
    setup();
    Report.generate();
    const sheet = Report.generate();
    T.assertEquals(1, sheet.getCharts().length, "Old charts should not accumulate.");
  });
})();
