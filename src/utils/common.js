/**
 * @fileoverview Common Utility Functions for Financial Planning Tools.
 * This file contains a collection of utility functions that are broadly used across
 * various modules and features of the Financial Planning Tools project.
 * These utilities are encapsulated within the `FinancialPlanner.Utils` namespace.
 * @module utils/common
 */

/**
 * @namespace FinancialPlanner.Utils
 * @description Provides common utility functions used across the Financial Planning Tools application,
 * such as column letter conversion, date formatting, sheet manipulation, and cell formatting.
 */
FinancialPlanner.Utils = (function() {
  /**
   * Array of English month names, used by `getMonthName`.
   * @private
   * @const {Array<string>}
   */
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
  
  // Public API
  return {
    /**
     * Converts a 1-based column index into its corresponding letter representation (e.g., 1 -> 'A', 27 -> 'AA').
     * @param {number} column - The 1-based column index. Must be a positive integer.
     * @return {string} The column letter(s). Returns an empty string if the input is not a positive number.
     * @memberof FinancialPlanner.Utils
     * @example
     * const colLetter = FinancialPlanner.Utils.columnToLetter(3); // Returns 'C'
     * const colLetter2 = FinancialPlanner.Utils.columnToLetter(28); // Returns 'AB'
     */
    columnToLetter: function(column) {
      let temp, letter = '';
      while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
      }
      return letter;
    },

    /**
     * Gets the full English name of a month from its 0-based index.
     * @param {number} monthIndex - The month index (0 for January, 1 for February, etc.).
     * @return {string | undefined} The full name of the month (e.g., "January"), or `undefined` for invalid indices.
     * @memberof FinancialPlanner.Utils
     * @example
     * const month = FinancialPlanner.Utils.getMonthName(0); // Returns 'January'
     */
    getMonthName: function(monthIndex) {
      return months[monthIndex];
    },

    /**
     * Retrieves a sheet by its name within the given spreadsheet.
     * If the sheet doesn't exist, it creates a new one with that name.
     * If the sheet *does* exist, its content (but not formatting or protected ranges) is cleared.
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object to operate on.
     * @param {string} sheetName - The desired name of the sheet.
     * @return {GoogleAppsScript.Spreadsheet.Sheet} The existing or newly created sheet object.
     * @memberof FinancialPlanner.Utils
     * @example
     * const ss = SpreadsheetApp.getActiveSpreadsheet();
     * const reportSheet = FinancialPlanner.Utils.getOrCreateSheet(ss, "Monthly Report");
     */
    getOrCreateSheet: function(spreadsheet, sheetName) {
      let sheet;
      try {
        sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          sheet = spreadsheet.insertSheet(sheetName);
        } else {
          // Clear existing content but preserve formatting
          sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
        }
      } catch (e) {
        sheet = spreadsheet.insertSheet(sheetName);
      }
      return sheet;
    },

    /**
     * Formats a given cell range as currency using a provided Google Sheets number format string.
     * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range to format.
     * @param {string} numberFormatString - The complete number format string to apply (e.g., "$#,##0.00;($#,##0.00)").
     * @return {GoogleAppsScript.Spreadsheet.Range} The same range object, allowing for method chaining.
     * @memberof FinancialPlanner.Utils
     * @example
     * const amountRange = sheet.getRange("C2:C10");
     * const formatStr = FinancialPlanner.Config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT;
     * FinancialPlanner.Utils.formatAsCurrency(amountRange, formatStr);
     */
    formatAsCurrency: function(range, numberFormatString) {
      // Using the provided Google Sheets format for currency
      range.setNumberFormat(numberFormatString);
      return range; // Return for chaining
    },

    /**
     * Formats a given cell range as a percentage with a specified number of decimal places.
     * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range to format.
     * @param {number} [decimalPlaces=1] - The number of decimal places to display (e.g., 1 for "0.0%", 2 for "0.00%").
     * @return {GoogleAppsScript.Spreadsheet.Range} The same range object, allowing for method chaining.
     * @memberof FinancialPlanner.Utils
     * @example
     * const rateRange = sheet.getRange("D2:D10");
     * FinancialPlanner.Utils.formatAsPercentage(rateRange, 2); // Format as 0.00%
     */
    formatAsPercentage: function(range, decimalPlaces = 1) {
      const format = `0.${'0'.repeat(decimalPlaces)}%`;
      range.setNumberFormat(format);
      return range; // Return for chaining
    },

    /**
     * Applies a background color to alternating rows within a specified range for improved readability (banding).
     * Starts applying the color from `startRow` and continues every other row up to `endRow`,
     * across all columns that have content in the sheet.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object to format.
     * @param {number} startRow - The 1-based starting row index for applying the alternating color.
     * @param {number} endRow - The 1-based ending row index for applying the alternating color.
     * @param {string} [color='#f9f9f9'] - The background color to apply (hex code or standard color name).
     * @return {GoogleAppsScript.Spreadsheet.Sheet} The same sheet object, allowing for method chaining.
     * @memberof FinancialPlanner.Utils
     * @example
     * FinancialPlanner.Utils.setAlternatingRowColors(reportSheet, 2, 20, '#eeeeee');
     */
    setAlternatingRowColors: function(sheet, startRow, endRow, color = '#f9f9f9') {
      for (let i = startRow; i <= endRow; i += 2) {
        sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground(color);
      }
      return sheet; // Return for chaining
    }
  };
})();

// For backward compatibility, create global references to the utility functions
// These can be removed once all code has been updated to use the namespace.

/**
 * Converts column index to letter. Delegates to `FinancialPlanner.Utils.columnToLetter`.
 * @param {number} column - The 1-based column index.
 * @return {string | undefined} The column letter(s) or undefined if the service isn't loaded.
 * @global
 */
function columnToLetter(column) {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Utils && FinancialPlanner.Utils.columnToLetter) {
    return FinancialPlanner.Utils.columnToLetter(column);
  }
  Logger.log("Global columnToLetter: FinancialPlanner.Utils not available.");
}

/**
 * Gets month name from index. Delegates to `FinancialPlanner.Utils.getMonthName`.
 * @param {number} monthIndex - The 0-based month index.
 * @return {string | undefined} The month name or undefined if the service isn't loaded.
 * @global
 */
function getMonthName(monthIndex) {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Utils && FinancialPlanner.Utils.getMonthName) {
    return FinancialPlanner.Utils.getMonthName(monthIndex);
  }
   Logger.log("Global getMonthName: FinancialPlanner.Utils not available.");
}

/**
 * Gets or creates a sheet by name. Delegates to `FinancialPlanner.Utils.getOrCreateSheet`.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object.
 * @param {string} sheetName - The name of the sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet | undefined} The sheet object or undefined if the service isn't loaded.
 * @global
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Utils && FinancialPlanner.Utils.getOrCreateSheet) {
    return FinancialPlanner.Utils.getOrCreateSheet(spreadsheet, sheetName);
  }
   Logger.log("Global getOrCreateSheet: FinancialPlanner.Utils not available.");
}

/**
 * Formats a range as currency using the default currency format from config. Delegates to `FinancialPlanner.Utils.formatAsCurrency`.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to format.
 * @return {GoogleAppsScript.Spreadsheet.Range | undefined} The range object or undefined if the service isn't loaded.
 * @global
 */
function formatAsCurrency(range) {
  if (typeof FinancialPlanner !== 'undefined' && 
      FinancialPlanner.Utils && 
      FinancialPlanner.Utils.formatAsCurrency &&
      FinancialPlanner.Config &&
      FinancialPlanner.Config.getLocale) {
    const defaultFormat = FinancialPlanner.Config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT;
    return FinancialPlanner.Utils.formatAsCurrency(range, defaultFormat);
  }
   Logger.log("Global formatAsCurrency: FinancialPlanner.Utils or FinancialPlanner.Config not available.");
}

/**
 * Formats a range as percentage. Delegates to `FinancialPlanner.Utils.formatAsPercentage`.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to format.
 * @param {number} [decimalPlaces=1] - The number of decimal places.
 * @return {GoogleAppsScript.Spreadsheet.Range | undefined} The range object or undefined if the service isn't loaded.
 * @global
 */
function formatAsPercentage(range, decimalPlaces = 1) {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Utils && FinancialPlanner.Utils.formatAsPercentage) {
    return FinancialPlanner.Utils.formatAsPercentage(range, decimalPlaces);
  }
   Logger.log("Global formatAsPercentage: FinancialPlanner.Utils not available.");
}

/**
 * Sets alternating row colors. Delegates to `FinancialPlanner.Utils.setAlternatingRowColors`.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} startRow - The starting row index.
 * @param {number} endRow - The ending row index.
 * @param {string} [color='#f9f9f9'] - The background color.
 * @return {GoogleAppsScript.Spreadsheet.Sheet | undefined} The sheet object or undefined if the service isn't loaded.
 * @global
 */
function setAlternatingRowColors(sheet, startRow, endRow, color = '#f9f9f9') {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Utils && FinancialPlanner.Utils.setAlternatingRowColors) {
    return FinancialPlanner.Utils.setAlternatingRowColors(sheet, startRow, endRow, color);
  }
   Logger.log("Global setAlternatingRowColors: FinancialPlanner.Utils not available.");
}
