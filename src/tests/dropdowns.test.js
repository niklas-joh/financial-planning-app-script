/**
 * Financial Planning Tools - Dropdown Service Tests
 *
 * Exercises the real FinancialPlanner.DropdownService against the in-memory
 * Apps Script fakes: a "Dropdowns" definition sheet drives the dependent
 * Type → Category → Sub-Category validations on the "Transactions" sheet.
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Dropdowns = FinancialPlanner.DropdownService;
  const PLACEHOLDER = 'Please select';
  const PENDING = '#eeeeee';
  const TYPE = 3, CATEGORY = 4, SUB = 5;

  let gas, tx, scriptCache;

  function setup() {
    gas = FakeGas.install(global);
    scriptCache = {};
    global.CacheService = {
      getScriptCache: () => ({
        get: (k) => (k in scriptCache ? scriptCache[k] : null),
        put: (k, v) => { scriptCache[k] = v; },
        remove: (k) => { delete scriptCache[k]; },
      }),
    };
    const defs = gas.spreadsheet.insertSheet('Dropdowns');
    defs.getRange(1, 1, 6, 3).setValues([
      ['Type', 'Category', 'Sub-Category'],
      ['Expense', 'Food', 'Groceries'],
      ['Expense', 'Food', 'Restaurants'],
      ['Expense', 'Housing', 'Rent'],
      ['Income', 'Salary', ''],
      ['Income', 'Bonus', ''],
    ]);
    tx = gas.spreadsheet.insertSheet('Transactions');
    tx.getRange(1, 1, 1, 6).setValues([['Date', 'Description', 'Type', 'Category', 'Sub-Category', 'Amount']]);
    tx.setFrozenRows(1);
    // Rebuild the service's in-memory cache from this test's definitions.
    Dropdowns.refreshCache();
  }

  function edit(row, col, value) {
    const range = tx.getRange(row, col);
    if (value !== undefined) range.setValue(value);
    Dropdowns.handleEdit({ range: range, source: gas.spreadsheet, value: value });
  }

  function listOptions(row, col) {
    const rule = tx.getRange(row, col).getDataValidation();
    return rule ? rule.values : null;
  }

  T.registerTest("DropdownService", "refreshCache should rebuild options and store them in the script cache", function() {
    setup();
    const cached = JSON.parse(scriptCache.dropdownsData);
    T.assertDeepEquals(['Food', 'Housing'], cached.typeToCategories.Expense, "Expense categories should be cached.");
    T.assertDeepEquals(['Groceries', 'Restaurants'], cached.typeCategoryToSubCategories['Expense___Food'], "Sub-categories should be keyed by type and category.");
    T.assertEquals('Dropdown options have been refreshed.', gas.spreadsheet.toasts.pop().message, "A success toast should be shown.");
  });

  T.registerTest("DropdownService", "initializeCache should not rebuild an existing cache", function() {
    setup();
    gas.spreadsheet.deleteSheet(gas.spreadsheet.getSheetByName('Dropdowns'));
    Dropdowns.initializeCache(); // must reuse the cache built in setup()
    edit(2, TYPE, 'Income');
    T.assertDeepEquals([PLACEHOLDER, 'Salary', 'Bonus'], listOptions(2, CATEGORY), "Existing cache should still be used.");
  });

  T.registerTest("DropdownService", "handleEdit on Type should update Category dropdown", function() {
    setup();
    edit(2, TYPE, 'Expense');
    T.assertDeepEquals([PLACEHOLDER, 'Food', 'Housing'], listOptions(2, CATEGORY), "Category list should match the type.");
    T.assertEquals(PENDING, tx.getRange(2, CATEGORY).getBackground(), "Empty category should be highlighted as pending.");
  });

  T.registerTest("DropdownService", "handleEdit on Category should update Sub-Category dropdown", function() {
    setup();
    edit(2, TYPE, 'Expense');
    edit(2, CATEGORY, 'Food');
    T.assertDeepEquals([PLACEHOLDER, 'Groceries', 'Restaurants'], listOptions(2, SUB), "Sub-category list should match type+category.");
    T.assertEquals(PENDING, tx.getRange(2, SUB).getBackground(), "Empty sub-category should be highlighted.");
    T.assertEquals('#ffffff', tx.getRange(2, CATEGORY).getBackground(), "Chosen category should no longer be highlighted.");
  });

  T.registerTest("DropdownService", "handleEdit on SubCategory should clear highlight", function() {
    setup();
    edit(2, TYPE, 'Expense');
    edit(2, CATEGORY, 'Food');
    edit(2, SUB, 'Groceries');
    T.assertEquals('#ffffff', tx.getRange(2, SUB).getBackground(), "Highlight should be cleared once chosen.");
    T.assertEquals('Groceries', tx.getRange(2, SUB).getValue(), "Chosen value should be kept.");
  });

  T.registerTest("DropdownService", "selecting the placeholder clears it", function() {
    setup();
    edit(2, TYPE, 'Expense');
    edit(2, CATEGORY, 'Food');
    edit(2, SUB, PLACEHOLDER);
    T.assertEquals('', tx.getRange(2, SUB).getValue(), "Placeholder text should not be stored.");
  });

  T.registerTest("DropdownService", "handleEdit changing Type should clear invalid Category/SubCategory", function() {
    setup();
    edit(2, TYPE, 'Expense');
    edit(2, CATEGORY, 'Food');
    edit(2, SUB, 'Groceries');
    edit(2, TYPE, 'Income');
    T.assertEquals('', tx.getRange(2, CATEGORY).getValue(), "Category should be cleared as 'Food' is not valid for 'Income'.");
    T.assertEquals('', tx.getRange(2, SUB).getValue(), "Sub-category should be cleared too.");
    T.assertTrue(listOptions(2, SUB) === null, "Sub-category validation should be removed.");
    T.assertDeepEquals([PLACEHOLDER, 'Salary', 'Bonus'], listOptions(2, CATEGORY), "Category list should switch to Income.");
  });

  T.registerTest("DropdownService", "pasting several rows updates each row", function() {
    setup();
    tx.getRange(2, TYPE, 2, 1).setValues([['Expense'], ['Income']]);
    Dropdowns.handleEdit({ range: tx.getRange(2, TYPE, 2, 1), source: gas.spreadsheet });
    T.assertDeepEquals([PLACEHOLDER, 'Food', 'Housing'], listOptions(2, CATEGORY), "Row 2 should get expense categories.");
    T.assertDeepEquals([PLACEHOLDER, 'Salary', 'Bonus'], listOptions(3, CATEGORY), "Row 3 should get income categories.");
  });

  T.registerTest("DropdownService", "edits on the header row or other sheets are ignored", function() {
    setup();
    Dropdowns.handleEdit({ range: tx.getRange(1, TYPE), source: gas.spreadsheet });
    T.assertTrue(listOptions(1, CATEGORY) === null, "Header row should not get validations.");
    const other = gas.spreadsheet.insertSheet('Other');
    other.getRange(2, TYPE).setValue('Expense');
    Dropdowns.handleEdit({ range: other.getRange(2, TYPE), source: gas.spreadsheet });
    T.assertTrue(other.getRange(2, CATEGORY).getDataValidation() === null, "Other sheets should be ignored.");
  });

  T.registerTest("DropdownService", "global onEdit routes Transactions edits to the dropdowns", function() {
    setup();
    tx.getRange(2, TYPE).setValue('Expense');
    global.onEdit({ range: tx.getRange(2, TYPE), source: gas.spreadsheet });
    T.assertDeepEquals([PLACEHOLDER, 'Food', 'Housing'], listOptions(2, CATEGORY), "The single global onEdit should reach DropdownService.");
  });
})();
