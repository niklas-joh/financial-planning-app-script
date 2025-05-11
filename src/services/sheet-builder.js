/**
 * @fileoverview Sheet Builder Service for Financial Planning Tools.
 * Provides a fluent API for programmatically constructing and formatting
 * Google Sheets. This service simplifies tasks like adding headers, data rows,
 * formulas, and applying various formatting options.
 * This module is designed to be instantiated by `00_module_loader.js`.
 * @module services/sheet-builder
 */

/**
 * IIFE to encapsulate the SheetBuilderModule logic.
 * @returns {function} The SheetBuilderModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const SheetBuilderModule = (function() {
  /**
   * Constructor for the SheetBuilderModule.
   * This module acts as a factory for creating `SheetBuilder` instances.
   * @param {ConfigModule} configInstance - An instance of ConfigModule.
   * @param {UtilsModule} utilsInstance - An instance of UtilsModule (assuming FinancialPlanner.Utils is UtilsModule).
   * @constructor
   * @alias SheetBuilderModule
   * @memberof module:services/sheet-builder
   */
  function SheetBuilderModuleConstructor(configInstance, utilsInstance) {
    /**
     * Instance of ConfigModule.
     * @type {ConfigModule}
     * @private
     */
    this.config = configInstance;
    /**
     * Instance of UtilsModule.
     * @type {UtilsModule}
     * @private
     */
    this.utils = utilsInstance;
  }

  /**
   * @classdesc Provides a fluent interface for building and formatting a Google Sheet.
   * Manages the current row and applies operations sequentially.
   * Instances are created via `SheetBuilderModule.create()`.
   * @class SheetBuilder
   * @private
   */
  class SheetBuilder {
    /**
     * Creates an instance of SheetBuilder.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheet object to be built upon.
     * @param {ConfigModule} config - An instance of the ConfigModule.
     * @param {UtilsModule} utils - An instance of the UtilsModule.
     */
    constructor(sheet, config, utils) {
      /**
       * The Google Sheet object being manipulated.
       * @type {GoogleAppsScript.Spreadsheet.Sheet}
       */
      this.sheet = sheet;
      /**
       * Instance of ConfigModule.
       * @type {ConfigModule}
       * @private
       */
      this.config = config;
      /**
       * Instance of UtilsModule.
       * @type {UtilsModule}
       * @private
       */
      this.utils = utils;
      /**
       * The current row number (1-based) where the next operation will start.
       * @type {number}
       */
      this.currentRow = 1;
      /**
       * Placeholder for future batch operations; not currently used.
       * @type {Array<object>}
       * @private
       */
      this.operations = [];
    }

    /**
     * Clears all content, formatting, and data validations from the sheet.
     * Resets the `currentRow` to 1.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    clear() {
      this.currentRow = 1; // Reset current row as sheet is cleared
      this.sheet.clear();
      this.sheet.clearFormats();
      this.sheet.getRange("A1:Z1000").setDataValidation(null);
      return this;
    }

    /**
     * Adds a header row to the sheet with specified values and formatting.
     * Increments the `currentRow` counter.
     * @param {Array<string>} headers - An array of strings for the header cells.
     * @param {object} [formatting={}] - Optional formatting object to apply to the header row.
     *   See `_applyFormatting` for available options.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    addHeaderRow(headers, formatting = {}) {
      const range = this.sheet.getRange(this.currentRow, 1, 1, headers.length);
      range.setValues([headers]);
      
      this._applyFormatting(range, formatting);
      this.currentRow++;
      return this;
    }

    /**
     * Adds a section header row, potentially merging cells.
     * Increments the `currentRow` counter.
     * @param {string} title - The title for the section header.
     * @param {{merge?: number, background?: string, fontColor?: string, fontWeight?: string, fontSize?: number, horizontalAlignment?: string}} [formatting={}] -
     *   Optional formatting object. If `formatting.merge` (number) is provided,
     *   the header cell will be merged across that many columns.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    addSectionHeader(title, formatting = {}) {
      const range = this.sheet.getRange(this.currentRow, 1);
      range.setValue(title);
      
      if (formatting.merge) {
        const mergeRange = this.sheet.getRange(this.currentRow, 1, 1, formatting.merge);
        mergeRange.merge();
      }
      
      this._applyFormatting(range, formatting);
      this.currentRow++;
      return this;
    }

    /**
     * Adds multiple rows of data to the sheet, with optional formulas and formatting.
     * Increments `currentRow` by the number of data rows added.
     * @param {Array<Array<*>>} data - A 2D array of data to write.
     * @param {{
     *   formulas?: Array<{startColumn: number, values: Array<Array<string>>}>,
     *   formatting?: object,
     *   rowFormatting?: Array<object|null>
     * }} [options={}] - Optional settings:
     *   - `formulas`: Array of formula configurations. Each config specifies `startColumn` (1-based)
     *     and `values` (a 2D array of formula strings, matching the dimensions of data rows for those columns).
     *   - `formatting`: A global formatting object to apply to all added data cells.
     *   - `rowFormatting`: An array of formatting objects, one for each data row.
     *     Null or undefined entries skip formatting for that row.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    addDataRows(data, options = {}) {
      if (data.length === 0) return this;
      
      const startRow = this.currentRow;
      const dataRange = this.sheet.getRange(
        this.currentRow, 1, data.length, data[0].length
      );
      dataRange.setValues(data);
      
      // Apply formulas if provided
      if (options.formulas && options.formulas.length > 0) {
        options.formulas.forEach(formulaConfig => {
          const formulaRange = this.sheet.getRange(
            startRow,
            formulaConfig.startColumn,
            data.length,
            formulaConfig.values[0].length
          );
          formulaRange.setFormulas(formulaConfig.values);
        });
      }
      
      // Apply formatting if provided
      if (options.formatting) {
        this._applyFormatting(dataRange, options.formatting);
      }
      
      // Apply row-specific formatting if provided
      if (options.rowFormatting) {
        options.rowFormatting.forEach((format, index) => {
          if (format) {
            const rowRange = this.sheet.getRange(
              startRow + index, 1, 1, data[0].length
            );
            this._applyFormatting(rowRange, format);
          }
        });
      }
      
      this.currentRow += data.length;
      return this;
    }

    /**
     * Adds a summary row with a label in the first column and specified formulas in other columns.
     * Increments the `currentRow` counter.
     * @param {string} label - The label for the summary row (placed in the first column).
     * @param {Array<{column: number, value: string}>} formulas - An array of objects, each specifying
     *   a 1-based `column` number and a formula `value` string.
     * @param {object} [formatting={}] - Optional formatting object to apply to the summary row.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    addSummaryRow(label, formulas, formatting = {}) {
      this.sheet.getRange(this.currentRow, 1).setValue(label);
      
      formulas.forEach(formula => {
        this.sheet.getRange(this.currentRow, formula.column).setFormula(formula.value);
      });
      
      if (formatting) {
        const lastColumn = formulas.length > 0 ? 
          Math.max(...formulas.map(f => f.column)) : 
          this.sheet.getLastColumn();
        
        const range = this.sheet.getRange(
          this.currentRow, 1, 1, lastColumn
        );
        this._applyFormatting(range, formatting);
      }
      
      this.currentRow++;
      return this;
    }

    /**
     * Adds a blank row for spacing by setting its height.
     * Increments the `currentRow` counter.
     * Note: This method sets the height of the current row to achieve a "blank" visual effect,
     * it does not insert new rows in the traditional sense if content follows.
     * @param {number} [height] - The height for the blank row in pixels.
     *   Defaults to the sheet's default row height if not specified.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    addBlankRow(height) {
      const effectiveHeight = height === undefined ? this.sheet.getDefaultRowHeight() : height;
      this.sheet.setRowHeight(this.currentRow, effectiveHeight);
      this.currentRow++;
      return this;
    }

    /**
     * Applies various formatting options to a given cell range.
     * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to format.
     * @param {{
     *   background?: string,
     *   fontColor?: string,
     *   fontWeight?: string,
     *   fontSize?: number,
     *   horizontalAlignment?: string,
     *   verticalAlignment?: string,
     *   numberFormat?: string,
     *   wrap?: boolean,
     *   border?: {top?: boolean, left?: boolean, bottom?: boolean, right?: boolean, internal?: boolean, color?: string, style?: GoogleAppsScript.Spreadsheet.BorderStyleValue},
     *   indent?: number
     * }} formatting - An object containing formatting properties.
     * @private
     */
    _applyFormatting(range, formatting) {
      if (formatting.background) range.setBackground(formatting.background);
      if (formatting.fontColor) range.setFontColor(formatting.fontColor);
      if (formatting.fontWeight) range.setFontWeight(formatting.fontWeight);
      if (formatting.fontSize) range.setFontSize(formatting.fontSize);
      if (formatting.horizontalAlignment) range.setHorizontalAlignment(formatting.horizontalAlignment);
      if (formatting.verticalAlignment) range.setVerticalAlignment(formatting.verticalAlignment);
      if (formatting.numberFormat) range.setNumberFormat(formatting.numberFormat);
      if (formatting.wrap) range.setWrap(formatting.wrap);
      
      if (formatting.border) {
        range.setBorder(
          true, true, true, true, formatting.border.internal || false, formatting.border.internal || false,
          formatting.borderColor || '#000000',
          formatting.borderStyle || SpreadsheetApp.BorderStyle.SOLID
        );
      }
      
      if (formatting.indent) range.setIndentationLevel(formatting.indent);
    }

    /**
     * Gets the current row number (1-based) where the next operation will begin.
     * @returns {number} The current row number.
     */
    getCurrentRow() {
      return this.currentRow;
    }

    /**
     * Sets the current row number (1-based) for subsequent operations.
     * @param {number} row - The row number to set as current.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    setCurrentRow(row) {
      this.currentRow = row;
      return this;
    }

    /**
     * Sets the widths of specified columns.
     * @param {Object<number, number>} widths - An object where keys are 1-based column indices
     *   and values are the desired widths in pixels.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    setColumnWidths(widths) {
      Object.entries(widths).forEach(([col, width]) => {
        this.sheet.setColumnWidth(parseInt(col), width);
      });
      return this;
    }

    /**
     * Sets the heights of specified rows.
     * @param {Object<number, number>} heights - An object where keys are 1-based row indices
     *   and values are the desired heights in pixels.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    setRowHeights(heights) {
      Object.entries(heights).forEach(([row, height]) => {
        this.sheet.setRowHeight(parseInt(row), height);
      });
      return this;
    }

    /**
     * Freezes a specified number of top rows in the sheet.
     * @param {number} numRows - The number of rows to freeze.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    freezeRows(numRows) {
      this.sheet.setFrozenRows(numRows);
      return this;
    }

    /**
     * Freezes a specified number of leftmost columns in the sheet.
     * @param {number} numColumns - The number of columns to freeze.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    freezeColumns(numColumns) {
      this.sheet.setFrozenColumns(numColumns);
      return this;
    }

    /**
     * Inserts checkboxes into a specified range of cells.
     * @param {number} row - The starting row number (1-based).
     * @param {number} column - The starting column number (1-based).
     * @param {number} [numRows=1] - The number of rows for the checkbox range.
     * @param {number} [numColumns=1] - The number of columns for the checkbox range.
     * @returns {SheetBuilder} The `SheetBuilder` instance for chaining.
     */
    addCheckboxes(row, column, numRows = 1, numColumns = 1) {
      const range = this.sheet.getRange(row, column, numRows, numColumns);
      range.insertCheckboxes();
      return this;
    }

    /**
     * Finalizes the sheet building process. Currently, this method primarily serves
     * to return the sheet object and the last row number used.
     * In future, it could execute batched operations if `this.operations` were used.
     * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number}}
     *   An object containing the modified sheet and the last row number written to (or set).
     */
    finalize() {
      return {
        sheet: this.sheet,
        lastRow: this.currentRow - 1
      };
    }
  }

  // Public API
  /**
   * Creates a new `SheetBuilder` instance for the given sheet.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheet object to be built upon.
   * @returns {SheetBuilder} A new instance of the `SheetBuilder` class.
   * @memberof SheetBuilderModule
   */
  SheetBuilderModuleConstructor.prototype.create = function(sheet) {
    return new SheetBuilder(sheet, this.config, this.utils);
  };

  return SheetBuilderModuleConstructor;
})();
