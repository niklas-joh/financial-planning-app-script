/**
 * Financial Planning Tools - Testing Framework
 *
 * This module provides a simple testing framework for the Financial Planning Tools project.
 * It allows for registration and execution of tests, along with basic assertion capabilities.
 */
var FinancialPlanner = FinancialPlanner || {};

FinancialPlanner.Testing = (function() {
  const tests = {}; // { moduleName: { testName: testFunction } }
  const results = {
    passed: 0,
    failed: 0,
    skipped: 0, // For future use
    details: [] // { module: string, test: string, status: 'passed'|'failed'|'skipped', message?: string, error?: Error }
  };

  /**
   * Resets the test results.
   * @private
   */
  function resetResults_() {
    results.passed = 0;
    results.failed = 0;
    results.skipped = 0;
    results.details = [];
  }

  /**
   * Logs a test result.
   * @param {string} moduleName - The name of the module being tested.
   * @param {string} testName - The name of the test.
   * @param {'passed'|'failed'|'skipped'} status - The status of the test.
   * @param {string} [message] - An optional message.
   * @param {Error} [error] - An optional error object for failed tests.
   * @private
   */
  function logResult_(moduleName, testName, status, message, error) {
    results.details.push({
      module: moduleName,
      test: testName,
      status: status,
      message: message,
      error: error ? { name: error.name, message: error.message, stack: error.stack } : undefined
    });
    if (status === 'passed') {
      results.passed++;
    } else if (status === 'failed') {
      results.failed++;
    } else if (status === 'skipped') {
      results.skipped++;
    }
  }

  // Public API
  return {
    /**
     * Registers a test function for a specific module.
     * @param {string} moduleName - The name of the module (e.g., "Utils", "ConfigService").
     * @param {string} testName - A descriptive name for the test (e.g., "should correctly format currency").
     * @param {function} testFunction - The function containing the test logic. It should throw an error if the test fails.
     */
    registerTest: function(moduleName, testName, testFunction) {
      if (!tests[moduleName]) {
        tests[moduleName] = {};
      }
      if (tests[moduleName][testName]) {
        console.warn(`Test "${testName}" already registered for module "${moduleName}". Overwriting.`);
      }
      tests[moduleName][testName] = testFunction;
    },

    /**
     * Runs all registered tests.
     * Logs results to the Google Apps Script Logger.
     * @return {Object} An object containing the test results summary.
     */
    runAll: function() {
      resetResults_();
      Logger.log("Starting test execution...");
      Logger.log("====================================");

      Object.keys(tests).forEach(moduleName => {
        Logger.log(`\n--- Testing Module: ${moduleName} ---`);
        const moduleTests = tests[moduleName];
        Object.keys(moduleTests).forEach(testName => {
          try {
            moduleTests[testName]();
            logResult_(moduleName, testName, 'passed', `✓ ${testName}`);
            Logger.log(`  ✓ ${testName}`);
          } catch (error) {
            logResult_(moduleName, testName, 'failed', `✗ ${testName} - ${error.message}`, error);
            Logger.log(`  ✗ ${testName} - FAILED: ${error.message}${error.stack ? `\n    Stack: ${error.stack}` : ''}`);
          }
        });
      });

      Logger.log("\n====================================");
      Logger.log("Test Execution Summary:");
      Logger.log(`  Passed: ${results.passed}`);
      Logger.log(`  Failed: ${results.failed}`);
      Logger.log(`  Skipped: ${results.skipped}`);
      Logger.log(`  Total: ${results.passed + results.failed + results.skipped}`);
      Logger.log("====================================");
      
      // For more detailed programmatic access if needed
      return { ...results }; 
    },

    /**
     * Runs all tests for a specific module.
     * @param {string} moduleName - The name of the module to test.
     * @return {Object} An object containing the test results summary for the module.
     */
    runModule: function(moduleName) {
      resetResults_();
      Logger.log(`Starting test execution for module: ${moduleName}...`);
      Logger.log("====================================");

      if (!tests[moduleName]) {
        Logger.log(`No tests found for module: ${moduleName}`);
        Logger.log("====================================");
        return { ...results };
      }

      const moduleTests = tests[moduleName];
      Object.keys(moduleTests).forEach(testName => {
        try {
          moduleTests[testName]();
          logResult_(moduleName, testName, 'passed', `✓ ${testName}`);
          Logger.log(`  ✓ ${testName}`);
        } catch (error) {
          logResult_(moduleName, testName, 'failed', `✗ ${testName} - ${error.message}`, error);
          Logger.log(`  ✗ ${testName} - FAILED: ${error.message}${error.stack ? `\n    Stack: ${error.stack}` : ''}`);
        }
      });

      Logger.log("\n====================================");
      Logger.log(`Test Execution Summary for ${moduleName}:`);
      Logger.log(`  Passed: ${results.passed}`);
      Logger.log(`  Failed: ${results.failed}`);
      Logger.log(`  Skipped: ${results.skipped}`);
      Logger.log(`  Total: ${results.passed + results.failed + results.skipped}`);
      Logger.log("====================================");
      return { ...results };
    },

    // --- Assertion Helpers ---

    /**
     * Asserts that two values are strictly equal (===).
     * @param {*} expected - The expected value.
     * @param {*} actual - The actual value.
     * @param {string} [message] - Optional message to display on failure.
     */
    assertEquals: function(expected, actual, message) {
      if (expected !== actual) {
        const defaultMessage = `Assertion Failed: Expected "${expected}" (type: ${typeof expected}) but got "${actual}" (type: ${typeof actual}).`;
        throw new Error(message || defaultMessage);
      }
    },

    /**
     * Asserts that a value is true.
     * @param {boolean} actual - The value to test.
     * @param {string} [message] - Optional message to display on failure.
     */
    assertTrue: function(actual, message) {
      if (actual !== true) {
        const defaultMessage = `Assertion Failed: Expected true but got ${actual}.`;
        throw new Error(message || defaultMessage);
      }
    },

    /**
     * Asserts that a value is false.
     * @param {boolean} actual - The value to test.
     * @param {string} [message] - Optional message to display on failure.
     */
    assertFalse: function(actual, message) {
      if (actual !== false) {
        const defaultMessage = `Assertion Failed: Expected false but got ${actual}.`;
        throw new Error(message || defaultMessage);
      }
    },

    /**
     * Asserts that a function throws an error.
     * @param {function} func - The function to execute.
     * @param {string|RegExp} [expectedError] - Optional. If a string, checks if the error message includes this string. If a RegExp, checks if the error message matches.
     * @param {string} [message] - Optional message to display on failure.
     */
    assertThrows: function(func, expectedError, message) {
      let caughtError = false;
      try {
        func();
      } catch (error) {
        caughtError = true;
        if (expectedError) {
          if (typeof expectedError === 'string' && !error.message.includes(expectedError)) {
            throw new Error(message || `Assertion Failed: Expected error message to include "${expectedError}", but got "${error.message}".`);
          } else if (expectedError instanceof RegExp && !expectedError.test(error.message)) {
            throw new Error(message || `Assertion Failed: Expected error message to match RegExp "${expectedError}", but got "${error.message}".`);
          }
        }
      }
      if (!caughtError) {
        throw new Error(message || "Assertion Failed: Expected function to throw an error, but it did not.");
      }
    },

    /**
     * Asserts that a value is not null or undefined.
     * @param {*} actual - The value to test.
     * @param {string} [message] - Optional message to display on failure.
     */
    assertNotNull: function(actual, message) {
      if (actual === null || actual === undefined) {
        const defaultMessage = `Assertion Failed: Expected value to be not null/undefined, but it was ${actual}.`;
        throw new Error(message || defaultMessage);
      }
    },
    
    /**
     * Asserts that two objects are deeply equal.
     * Note: This is a simple implementation. For complex objects or specific needs, a more robust deep equal library might be better.
     * @param {Object} expected - The expected object.
     * @param {Object} actual - The actual object.
     * @param {string} [message] - Optional message to display on failure.
     */
    assertDeepEquals: function(expected, actual, message) {
      try {
        const expectedJSON = JSON.stringify(expected);
        const actualJSON = JSON.stringify(actual);
        if (expectedJSON !== actualJSON) {
          throw new Error(message || `Assertion Failed: Expected deep equality. Expected: ${expectedJSON}, Actual: ${actualJSON}`);
        }
      } catch (e) {
         throw new Error(message || `Assertion Failed: Could not compare objects. Ensure they are JSON serializable. Error: ${e.message}`);
      }
    }
  };
})();

// Example of how to register a test (will be in separate test files)
/*
FinancialPlanner.Testing.registerTest("ExampleModule", "shouldAddTwoNumbers", function() {
  const result = 1 + 1;
  FinancialPlanner.Testing.assertEquals(2, result, "1 + 1 should be 2");
});

FinancialPlanner.Testing.registerTest("ExampleModule", "shouldBeTrue", function() {
  FinancialPlanner.Testing.assertTrue(true, "This should be true");
});
*/

// To run tests from a Google Apps Script function:
/*
function runAllTests() {
  FinancialPlanner.Testing.runAll();
}

function runUtilsTests() {
  FinancialPlanner.Testing.runModule("Utils");
}
*/
