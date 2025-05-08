/**
 * Financial Planning Tools - Config Module Tests
 *
 * This file contains tests for the FinancialPlanner.Config module.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;
  const Config = FinancialPlanner.Config;

  // --- Test Suite for FinancialPlanner.Config ---

  T.registerTest("Config", "should return the complete configuration object", function() {
    const config = Config.get();
    T.assertNotNull(config, "Config object should not be null.");
    T.assertTrue(typeof config === 'object', "Config should be an object.");
    T.assertNotNull(config.SHEETS, "Config should contain SHEETS section.");
  });

  T.registerTest("Config", "should return a specific configuration section", function() {
    const sheetsSection = Config.getSection('SHEETS');
    T.assertNotNull(sheetsSection, "SHEETS section should not be null.");
    T.assertEquals("Overview", sheetsSection.OVERVIEW, "SHEETS.OVERVIEW should be 'Overview'.");

    const colorsSection = Config.getSection('COLORS');
    T.assertNotNull(colorsSection, "COLORS section should not be null.");
    T.assertNotNull(colorsSection.UI, "COLORS.UI section should exist.");
    T.assertEquals("#C62828", colorsSection.UI.HEADER_BG, "COLORS.UI.HEADER_BG should be '#C62828'.");
  });

  T.registerTest("Config", "should return an empty object for non-existent section", function() {
    const nonExistent = Config.getSection('NON_EXISTENT_SECTION');
    T.assertNotNull(nonExistent, "Result for non-existent section should not be null.");
    T.assertTrue(typeof nonExistent === 'object', "Result should be an object.");
    T.assertEquals(0, Object.keys(nonExistent).length, "Result for non-existent section should be an empty object.");
  });

  T.registerTest("Config", "getSheetNames should return the SHEETS section", function() {
    const sheetNames = Config.getSheetNames();
    const sheetsSection = Config.getSection('SHEETS');
    T.assertDeepEquals(sheetsSection, sheetNames, "getSheetNames() should return the same as getSection('SHEETS').");
    T.assertEquals("Transactions", sheetNames.TRANSACTIONS, "Sheet name for Transactions should be correct.");
  });

  T.registerTest("Config", "getTransactionTypes should return the TRANSACTION_TYPES section", function() {
    const transactionTypes = Config.getTransactionTypes();
    const transactionTypesSection = Config.getSection('TRANSACTION_TYPES');
    T.assertDeepEquals(transactionTypesSection, transactionTypes, "getTransactionTypes() should return the same as getSection('TRANSACTION_TYPES').");
    T.assertEquals("Income", transactionTypes.INCOME, "Transaction type for Income should be correct.");
  });

  T.registerTest("Config", "getLocale should return the LOCALE section", function() {
    const locale = Config.getLocale();
    const localeSection = Config.getSection('LOCALE');
    T.assertDeepEquals(localeSection, locale, "getLocale() should return the same as getSection('LOCALE').");
    T.assertEquals("€", locale.CURRENCY_SYMBOL, "Currency symbol should be correct.");
  });

  // Note: update() and reset() modify internal state, which isn't directly exposed.
  // Testing them relies on observing changes via get() or getSection().
  // This test assumes update() is not meant for persistent changes across script executions.
  T.registerTest("Config", "update and reset should modify/revert config temporarily", function() {
    const originalSymbol = Config.getLocale().CURRENCY_SYMBOL;
    const newSymbol = "$";

    try {
      // Update
      Config.update({ LOCALE: { CURRENCY_SYMBOL: newSymbol } });
      const updatedLocale = Config.getLocale();
      T.assertEquals(newSymbol, updatedLocale.CURRENCY_SYMBOL, `After update, currency symbol should be '${newSymbol}'.`);

      // Ensure other sections are not affected if not updated (simple check)
      T.assertEquals("Overview", Config.getSheetNames().OVERVIEW, "Other sections should remain unchanged after update.");

      // Reset
      Config.reset();
      const resetLocale = Config.getLocale();
      T.assertEquals(originalSymbol, resetLocale.CURRENCY_SYMBOL, `After reset, currency symbol should be '${originalSymbol}'.`);

    } finally {
      // Ensure reset happens even if assertions fail
      Config.reset();
    }
  });

  T.registerTest("Config", "get should return a deep copy", function() {
    const config1 = Config.get();
    const config2 = Config.get();

    T.assertFalse(config1 === config2, "Two calls to get() should return different object references.");
    T.assertDeepEquals(config1, config2, "The content of the two config objects should be the same initially.");

    // Modify one copy and check the other is unaffected
    config1.SHEETS.OVERVIEW = "MODIFIED_OVERVIEW";
    T.assertEquals("MODIFIED_OVERVIEW", config1.SHEETS.OVERVIEW, "Modified copy should reflect the change.");
    T.assertFalse(config2.SHEETS.OVERVIEW === "MODIFIED_OVERVIEW", "Original config object (from second get()) should not be affected by modifying the first copy.");
    
    // Verify the actual service still returns the original value after modification of a copy
    const config3 = Config.get();
     T.assertFalse(config3.SHEETS.OVERVIEW === "MODIFIED_OVERVIEW", "A new call to get() should return the original unmodified value.");
  });


})(); // End IIFE
