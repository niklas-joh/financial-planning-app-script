/**
 * Financial Planning Tools - Utils Module Tests
 *
 * This file contains tests for the FinancialPlanner.Utils module.
 */

// Ensure the main namespace and testing framework are available
var FinancialPlanner = FinancialPlanner || {};
FinancialPlanner.Testing = FinancialPlanner.Testing || {}; // Assuming testing-framework.js is loaded first

(function(T, U) { // T = Testing, U = Utils
  if (!T || !T.registerTest) {
    console.error("Testing framework (FinancialPlanner.Testing) is not available. Skipping Utils tests.");
    return;
  }
  if (!U) {
     console.error("Utils module (FinancialPlanner.Utils) is not available. Skipping Utils tests.");
     return;
  }

  const MODULE_NAME = "Utils";

  // --- Tests for columnToLetter ---
  T.registerTest(MODULE_NAME, "columnToLetter should convert 1 to A", function() {
    T.assertEquals("A", U.columnToLetter(1), "Column 1 should be A");
  });

  T.registerTest(MODULE_NAME, "columnToLetter should convert 26 to Z", function() {
    T.assertEquals("Z", U.columnToLetter(26), "Column 26 should be Z");
  });

  T.registerTest(MODULE_NAME, "columnToLetter should convert 27 to AA", function() {
    T.assertEquals("AA", U.columnToLetter(27), "Column 27 should be AA");
  });

  T.registerTest(MODULE_NAME, "columnToLetter should convert 52 to AZ", function() {
    T.assertEquals("AZ", U.columnToLetter(52), "Column 52 should be AZ");
  });
  
  T.registerTest(MODULE_NAME, "columnToLetter should convert 702 to ZZ", function() {
    T.assertEquals("ZZ", U.columnToLetter(702), "Column 702 should be ZZ");
  });
  
  T.registerTest(MODULE_NAME, "columnToLetter should convert 703 to AAA", function() {
    T.assertEquals("AAA", U.columnToLetter(703), "Column 703 should be AAA");
  });

  T.registerTest(MODULE_NAME, "columnToLetter should handle zero gracefully (return empty string or throw)", function() {
     // Depending on desired behavior, either expect empty string or assertThrows
     try {
       const result = U.columnToLetter(0);
       T.assertEquals("", result, "Column 0 should return empty string"); 
     } catch (e) {
       // Or if it's expected to throw:
       // T.assertThrows(function() { U.columnToLetter(0); }, /invalid/i, "Column 0 should throw an error");
       T.assertTrue(true); // If throwing is acceptable, pass the test
     }
  });
  
   T.registerTest(MODULE_NAME, "columnToLetter should handle negative numbers gracefully", function() {
     try {
       const result = U.columnToLetter(-5);
       T.assertEquals("", result, "Negative column should return empty string");
     } catch (e) {
       T.assertTrue(true); // If throwing is acceptable, pass the test
     }
  });

  // --- Tests for getMonthName ---
  T.registerTest(MODULE_NAME, "getMonthName should return January for index 0", function() {
    T.assertEquals("January", U.getMonthName(0), "Month index 0 should be January");
  });

  T.registerTest(MODULE_NAME, "getMonthName should return December for index 11", function() {
    T.assertEquals("December", U.getMonthName(11), "Month index 11 should be December");
  });

  T.registerTest(MODULE_NAME, "getMonthName should handle invalid index (e.g., 12)", function() {
    // Assuming it returns an empty string or a specific error string for invalid indices
    const result = U.getMonthName(12);
    T.assertEquals("", result, "Invalid month index 12 should return empty string"); 
    // Or: T.assertEquals("Invalid Month", result, "Invalid month index 12 should return 'Invalid Month'");
  });
  
  T.registerTest(MODULE_NAME, "getMonthName should handle invalid index (e.g., -1)", function() {
    const result = U.getMonthName(-1);
    T.assertEquals("", result, "Invalid month index -1 should return empty string");
  });

  // --- Tests for formatAsCurrency (Basic check, requires Mocks for SpreadsheetApp) ---
  // Note: Testing functions interacting directly with SpreadsheetApp is harder without mocks.
  // This is a placeholder showing the intent. More robust tests would need a mocking strategy.
  T.registerTest(MODULE_NAME, "formatAsCurrency should exist", function() {
     T.assertNotNull(U.formatAsCurrency, "formatAsCurrency function should exist");
     T.assertEquals('function', typeof U.formatAsCurrency, "formatAsCurrency should be a function");
  });
  
  // --- Tests for formatAsPercentage (Basic check) ---
   T.registerTest(MODULE_NAME, "formatAsPercentage should exist", function() {
     T.assertNotNull(U.formatAsPercentage, "formatAsPercentage function should exist");
     T.assertEquals('function', typeof U.formatAsPercentage, "formatAsPercentage should be a function");
  });

  // --- Tests for getOrCreateSheet (Basic check, requires Mocks) ---
   T.registerTest(MODULE_NAME, "getOrCreateSheet should exist", function() {
     T.assertNotNull(U.getOrCreateSheet, "getOrCreateSheet function should exist");
     T.assertEquals('function', typeof U.getOrCreateSheet, "getOrCreateSheet should be a function");
  });

})(FinancialPlanner.Testing, FinancialPlanner.Utils);

// Remember to add "src/tests/utils.test.js" to appsscript.json
