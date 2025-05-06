/**
 * Financial Planning Tools - Common Utilities
 * 
 * This file contains utility functions that are used across multiple features
 * of the Financial Planning Tools project.
 */

/**
 * Converts column index to letter (e.g., 1 -> A, 27 -> AA)
 * @param {Number} column - The column index (1-based)
 * @return {String} The column letter
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Gets month name from index (0-11)
 * @param {Number} monthIndex - The month index (0-based, where 0 = January)
 * @return {String} The month name
 */
function getMonthName(monthIndex) {
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
  return months[monthIndex];
}

/**
 * Creates or gets a sheet by name, clearing its content if it already exists
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The spreadsheet to work with
 * @param {String} sheetName - The name of the sheet to create or get
 * @return {SpreadsheetApp.Sheet} The sheet
 */
function getOrCreateSheet(spreadsheet, sheetName) {
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
}

/**
 * Formats a range as currency
 * @param {SpreadsheetApp.Range} range - The range to format
 * @param {String} currencySymbol - The currency symbol to use (default: €)
 * @param {String} locale - The locale code for the currency (default: 2 for Euro)
 */
function formatAsCurrency(range, currencySymbol = '€', locale = '2') {
  // Using the specified Google Sheets format for currency
  range.setNumberFormat(`_-[$${currencySymbol}-${locale}]\\ * #,##0_-;\\-[$${currencySymbol}-${locale}]\\ * #,##0_-;_-[$${currencySymbol}-${locale}]\\ * "-"??_-;_-@`);
}

/**
 * Formats a range as percentage
 * @param {SpreadsheetApp.Range} range - The range to format
 * @param {Number} decimalPlaces - The number of decimal places to show (default: 1)
 */
function formatAsPercentage(range, decimalPlaces = 1) {
  const format = `0.${'0'.repeat(decimalPlaces)}%`;
  range.setNumberFormat(format);
}

/**
 * Sets alternating row colors for better readability
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to format
 * @param {Number} startRow - The row to start formatting from
 * @param {Number} endRow - The row to end formatting at
 * @param {String} color - The color to use for alternating rows (default: #f9f9f9)
 */
function setAlternatingRowColors(sheet, startRow, endRow, color = '#f9f9f9') {
  for (let i = startRow; i <= endRow; i += 2) {
    sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground(color);
  }
}
