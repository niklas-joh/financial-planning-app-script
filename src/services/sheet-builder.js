/**
 * @fileoverview Sheet Builder Service for Financial Planning Tools.
 * Provides a fluent API for building complex sheets.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const SheetBuilderModule = (function() {
  /**
   * Constructor for the SheetBuilderModule.
   * @param {object} configInstance - An instance of ConfigModule.
   * @param {object} utilsInstance - An instance of Utils.
   * @constructor
   */
  function SheetBuilderModuleConstructor(configInstance, utilsInstance) {
    this.config = configInstance;
    this.utils = utilsInstance;
  }

  /**
   * Internal SheetBuilder class
   */
  class SheetBuilder {
    constructor(sheet, config, utils) {
      this.sheet = sheet;
      this.config = config;
      this.utils = utils;
      this.currentRow = 1;
      this.operations = [];
    }

    /**
     * Clears the sheet and resets formatting
     */
    clear() {
      this.sheet.clear();
      this.sheet.clearFormats();
      this.sheet.getRange("A1:Z1000").setDataValidation(null);
      return this;
    }

    /**
     * Adds a header row with formatting
     */
    addHeaderRow(headers, formatting = {}) {
      const range = this.sheet.getRange(this.currentRow, 1, 1, headers.length);
      range.setValues([headers]);
      
      this._applyFormatting(range, formatting);
      this.currentRow++;
      return this;
    }

    /**
     * Adds a section header
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
     * Adds data rows with optional formulas
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
     * Adds a summary row
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
     * Adds a blank row for spacing
     */
    addBlankRow(height = 1) {
      this.sheet.setRowHeight(this.currentRow, height);
      this.currentRow++;
      return this;
    }

    /**
     * Applies formatting to a range
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
     * Gets the current row number
     */
    getCurrentRow() {
      return this.currentRow;
    }

    /**
     * Sets the current row number
     */
    setCurrentRow(row) {
      this.currentRow = row;
      return this;
    }

    /**
     * Sets column widths
     */
    setColumnWidths(widths) {
      Object.entries(widths).forEach(([col, width]) => {
        this.sheet.setColumnWidth(parseInt(col), width);
      });
      return this;
    }

    /**
     * Sets row heights
     */
    setRowHeights(heights) {
      Object.entries(heights).forEach(([row, height]) => {
        this.sheet.setRowHeight(parseInt(row), height);
      });
      return this;
    }

    /**
     * Freezes rows
     */
    freezeRows(numRows) {
      this.sheet.setFrozenRows(numRows);
      return this;
    }

    /**
     * Freezes columns
     */
    freezeColumns(numColumns) {
      this.sheet.setFrozenColumns(numColumns);
      return this;
    }

    /**
     * Adds checkboxes to a range
     */
    addCheckboxes(row, column, numRows = 1, numColumns = 1) {
      const range = this.sheet.getRange(row, column, numRows, numColumns);
      range.insertCheckboxes();
      return this;
    }

    /**
     * Finalizes the sheet with any remaining operations
     */
    finalize() {
      return {
        sheet: this.sheet,
        lastRow: this.currentRow - 1
      };
    }
  }

  // Public API
  SheetBuilderModuleConstructor.prototype.create = function(sheet) {
    return new SheetBuilder(sheet, this.config, this.utils);
  };

  return SheetBuilderModuleConstructor;
})();