/**
 * Financial Planning Tools - Common Utilities
 * 
 * This file contains utility functions that are used across multiple features
 * of the Financial Planning Tools project. These utilities are encapsulated
 * within the FinancialPlanner.Utils namespace to prevent global namespace pollution.
 */

/**
 * @namespace FinancialPlanner.Utils
 * @description Provides common utility functions used across the Financial Planning Tools application.
 */
FinancialPlanner.Utils = (function() {
  // Private variables and functions can be defined here
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
  
  // Public API
  return {
    /**
     * Converts a 1-based column index into its corresponding letter representation (e.g., 1 -> 'A', 27 -> 'AA').
     * @param {number} column - The 1-based column index.
     * @return {string} The column letter(s). Returns an empty string for non-positive indices.
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
     * @return {string} The full name of the month (e.g., "January"). Returns undefined for invalid indices.
     * @example
     * const month = FinancialPlanner.Utils.getMonthName(0); // Returns 'January'
     */
    getMonthName: function(monthIndex) {
      return months[monthIndex];
    },

    /**
     * Retrieves a sheet by its name within the given spreadsheet. If the sheet doesn't exist,
     * it creates a new one with that name. If the sheet *does* exist, its content (but not formatting) is cleared.
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object to operate on.
     * @param {string} sheetName - The desired name of the sheet.
     * @return {GoogleAppsScript.Spreadsheet.Sheet} The existing or newly created sheet object.
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
     * Formats a given cell range as currency using a specific Google Sheets number format string.
     * Allows customization of the currency symbol and locale identifier used in the format.
     * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range to format.
     * @param {string} [currencySymbol='€'] - The currency symbol to display (e.g., '$', '£').
     * @param {string} [locale='2'] - The locale identifier used in the format string (e.g., '1' for USD, '2' for EUR).
     *                                Note: This is specific to the Google Sheets format string structure.
     * @return {GoogleAppsScript.Spreadsheet.Range} The same range object, allowing for method chaining.
     * @example
     * const amountRange = sheet.getRange("C2:C10");
     * FinancialPlanner.Utils.formatAsCurrency(amountRange, '$', '1');
     */
    formatAsCurrency: function(range, currencySymbol = '€', locale = '2') {
      // Using the specified Google Sheets format for currency
      range.setNumberFormat(`_-[$${currencySymbol}-${locale}]\\ * #,##0_-;\\-[$${currencySymbol}-${locale}]\\ * #,##0_-;_-[$${currencySymbol}-${locale}]\\ * "-"??_-;_-@`);
      return range; // Return for chaining
    },

    /**
     * Formats a given cell range as a percentage with a specified number of decimal places.
     * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range to format.
     * @param {number} [decimalPlaces=1] - The number of decimal places to display after the percentage point.
     * @return {GoogleAppsScript.Spreadsheet.Range} The same range object, allowing for method chaining.
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
     * Starts applying the color from `startRow` and continues every other row up to `endRow`.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object to format.
     * @param {number} startRow - The 1-based starting row index for applying the color.
     * @param {number} endRow - The 1-based ending row index for applying the color.
     * @param {string} [color='#f9f9f9'] - The background color to apply (hex code or color name).
     * @return {GoogleAppsScript.Spreadsheet.Sheet} The same sheet object, allowing for method chaining.
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
 * Formats a range as currency. Delegates to `FinancialPlanner.Utils.formatAsCurrency`.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to format.
 * @param {string} [currencySymbol='€'] - The currency symbol.
 * @param {string} [locale='2'] - The locale identifier.
 * @return {GoogleAppsScript.Spreadsheet.Range | undefined} The range object or undefined if the service isn't loaded.
 * @global
 */
function formatAsCurrency(range, currencySymbol = '€', locale = '2') {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Utils && FinancialPlanner.Utils.formatAsCurrency) {
    return FinancialPlanner.Utils.formatAsCurrency(range, currencySymbol, locale);
  }
   Logger.log("Global formatAsCurrency: FinancialPlanner.Utils not available.");
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
