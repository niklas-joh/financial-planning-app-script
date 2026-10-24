/**
 * Financial Planning Tools - Finance Overview Tests
 *
 * Exercises the real FinancialPlanner.FinanceOverview end to end (with the
 * real DataProcessor, SheetBuilder, FormulaBuilder, CacheService, Settings
 * and FinancialAnalysisService) against the in-memory Apps Script fakes.
 * The overview is formula-driven, so the tests assert on the layout and the
 * formulas written, not on computed values.
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Overview = FinancialPlanner.FinanceOverview;
  let gas;

  function setup() {
    gas = FakeGas.install(global);
    const rows = [
      ['Date', 'Description', 'Type', 'Category', 'Sub-Category', 'Amount'],
      [new Date(2024, 0, 5), 'Employer', 'Income', 'Salary', '', 4000],
      [new Date(2024, 0, 6), 'Landlord', 'Essentials', 'Housing', 'Rent', -1500],
      [new Date(2024, 0, 7), 'Supermarket', 'Essentials', 'Food', 'Groceries', -300],
      [new Date(2024, 0, 8), 'Cinema', 'Wants/Pleasure', 'Fun', 'Cinema', -40],
      [new Date(2024, 0, 9), 'ETF', 'Savings', 'Investments', 'ETF', -500],
    ];
    gas.spreadsheet.insertSheet('Transactions').getRange(1, 1, rows.length, 6).setValues(rows);
  }

  function overview() {
    return gas.spreadsheet.getSheetByName('Overview');
  }

  function rowOf(label, column) {
    const values = overview()._values();
    for (let i = 0; i < values.length; i++) if (values[i][column || 0] === label) return i + 1;
    return -1;
  }

  T.registerTest("FinanceOverview", "create should initialize, process, build, and format", function() {
    setup();
    const result = Overview.create();
    T.assertNotNull(result, "create() should return the builder result.");
    const sheet = overview();
    T.assertNotNull(sheet, "Overview sheet should be created.");
    T.assertDeepEquals(FinancialPlanner.Config.getSection('HEADERS'), sheet.getRange(1, 1, 1, 18).getValues()[0], "Header row should match config.");
    T.assertEquals(1, sheet.getFrozenRows(), "Header should be frozen.");
    ['Income', 'Total Income', 'Expenses', 'Essentials', 'Wants/Pleasure', 'Total Expenses', 'Savings', 'Net Calculations'].forEach(function(label) {
      T.assertTrue(rowOf(label) > 0, "Overview should contain a '" + label + "' row.");
    });
    T.assertTrue(rowOf('Net (Total Income - Expenses - Savings)') > 0, "Net rows should be added when savings exist.");
    T.assertEquals('Financial overview generated successfully!', gas.spreadsheet.toasts.pop().message, "Success toast should be shown.");
  });

  T.registerTest("FinanceOverview", "category rows should get SUMIFS formulas against the Transactions sheet", function() {
    setup();
    Overview.create();
    const row = rowOf('Housing', 1);
    T.assertTrue(row > 0, "Housing row should exist.");
    const jan = overview().getRange(row, 5).getFormula();
    T.assertTrue(jan.indexOf('SUMIFS(Transactions!F:F') !== -1, "January should sum the Amount column (F): " + jan);
    T.assertTrue(jan.indexOf('/IF(D' + row + '=TRUE, 2, 1)') !== -1, "Expense rows should halve shared expenses.");
    T.assertEquals('=SUM(E' + row + ':P' + row + ')', overview().getRange(row, 17).getFormula().replace(/\s/g, ''), "Total column should sum the months.");
    T.assertEquals(false, overview().getRange(row, 4).getValue(), "Expense rows should get a 'Shared?' checkbox.");
  });

  T.registerTest("FinanceOverview", "create should also build the Analysis sheet", function() {
    setup();
    Overview.create();
    T.assertNotNull(gas.spreadsheet.getSheetByName('Analysis'), "Financial analysis should run after the overview.");
  });

  T.registerTest("FinanceOverview", "sub-category toggle is written from settings", function() {
    setup();
    Overview.create();
    const toggle = FinancialPlanner.Config.getSection('UI').SUBCATEGORY_TOGGLE;
    T.assertEquals(true, overview().getRange(toggle.CHECKBOX_CELL).getValue(), "Checkbox should reflect the default (show sub-categories).");
    T.assertEquals(toggle.LABEL_TEXT, overview().getRange(toggle.LABEL_CELL).getValue(), "Label should be written.");
  });

  T.registerTest("FinanceOverview", "handleEdit should update preference and recreate on checkbox change", function() {
    setup();
    Overview.create();
    const toggle = FinancialPlanner.Config.getSection('UI').SUBCATEGORY_TOGGLE;
    const checkbox = overview().getRange(toggle.CHECKBOX_CELL);
    checkbox.setValue(false);
    Overview.handleEdit({ range: checkbox, source: gas.spreadsheet });
    T.assertFalse(FinancialPlanner.SettingsService.getShowSubCategories(), "Preference should be saved.");
    T.assertEquals(false, overview().getRange(toggle.CHECKBOX_CELL).getValue(), "Rebuilt overview should keep the new setting.");
  });

  T.registerTest("FinanceOverview", "handleEdit should do nothing if edit is not on checkbox", function() {
    setup();
    Overview.create();
    const before = gas.spreadsheet.toasts.length;
    Overview.handleEdit({ range: overview().getRange('B5'), source: gas.spreadsheet });
    T.assertEquals(before, gas.spreadsheet.toasts.length, "No rebuild should start.");
    T.assertTrue(FinancialPlanner.SettingsService.getShowSubCategories(), "Preference should be unchanged.");
  });

  T.registerTest("FinanceOverview", "create should report a missing Transactions sheet", function() {
    gas = FakeGas.install(global);
    T.assertThrows(function() { Overview.create(); }, undefined, "create() re-throws so callers know it failed.");
    T.assertTrue(gas.ui.alerts[0].message.indexOf('Failed to generate financial overview') !== -1, "User should see the failure.");
  });
})();
