/**
 * @fileoverview Formula Builder Service for Financial Planning Tools.
 * Centralizes the logic for constructing various spreadsheet formulas,
 * making it easier to manage and update them across the application.
 * This module is designed to be instantiated by `00_module_loader.js`.
 * @module services/formula-builder
 */

/**
 * IIFE to encapsulate the FormulaBuilderModule logic.
 * @returns {function} The FormulaBuilderModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const FormulaBuilderModule = (function() {
  /**
   * Constructor for the FormulaBuilderModule.
   * @param {ConfigModule} configInstance - An instance of ConfigModule.
   * @constructor
   * @alias FormulaBuilderModule
   * @memberof module:services/formula-builder
   */
  function FormulaBuilderModuleConstructor(configInstance) {
    /**
     * Instance of ConfigModule.
     * @type {ConfigModule}
     * @private
     */
    this.config = configInstance;
  }

  // Private helper methods
  /**
   * Builds a standard criteria string part for SUMIFS/COUNTIFS.
   * @param {string} criteriaRange - The range for the criteria (e.g., "A1:A10").
   * @param {string|number} criteriaValue - The value for the criteria (e.g., "\"Apples\"" or "10").
   * @returns {string} The formatted criteria string (e.g., "A1:A10,\"Apples\"").
   * @private
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype._buildCriteriaString = function(criteriaRange, criteriaValue) {
    return `${criteriaRange},${criteriaValue}`;
  };

  /**
   * Builds a date-based criteria string part for SUMIFS/COUNTIFS.
   * @param {string} dateRange - The range containing dates.
   * @param {string} operator - The comparison operator (e.g., ">=", "<=").
   * @param {string} dateValue - The date value, typically formatted as a string recognized by Sheets.
   * @returns {string} The formatted date criteria string (e.g., "B1:B10,\">="&DATE(2023,1,1)\"").
   * @private
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype._buildDateCriteriaString = function(dateRange, operator, dateValue) {
    // Ensure dateValue is quoted if it's not a cell reference or function call
    const val = /^[A-Z]+\d*|^[A-Z]+\(/.test(dateValue) ? dateValue : `"${dateValue}"`;
    return `${dateRange},"${operator}"&${val}`;
  };

  /**
   * Builds a criteria string part with an operator for SUMIFS/COUNTIFS.
   * @param {string} criteriaRange - The range for the criteria.
   * @param {string} operator - The comparison operator (e.g., "<>", ">").
   * @param {string|number} criteriaValue - The value for the criteria.
   * @returns {string} The formatted criteria string (e.g., "C1:C10,\"<>0\"").
   * @private
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype._buildCriteriaOperatorString = function(criteriaRange, operator, criteriaValue) {
    // Ensure criteriaValue is quoted if it's text and not a cell reference
    const val = (typeof criteriaValue === 'string' && !/^[A-Z]+\d*/.test(criteriaValue) && !/^"[^"]*"$/.test(criteriaValue))
                ? `"${criteriaValue}"`
                : criteriaValue;
    return `${criteriaRange},"${operator}"&${val}`;
  };

  // Public API methods

  /**
   * Builds a SUMIFS formula, commonly used for monthly calculations.
   * Can optionally include a divisor for shared amounts.
   * @param {{sumRange: string, criteria: Array<{range: string, value: string|number, operator?: string, type?: 'date'}>, sharedDivisor?: string}} params - Parameters for the formula.
   *   - `sumRange`: The range of cells to sum (e.g., "C1:C100").
   *   - `criteria`: An array of criteria objects. Each object should have:
   *     - `range`: The criteria range (e.g., "A1:A100").
   *     - `value`: The criteria value (e.g., "\"Income\"", "B1").
   *     - `operator`: (Optional) The comparison operator (e.g., ">=", "<>").
   *     - `type`: (Optional) If 'date', special handling for date criteria is applied.
   *   - `sharedDivisor`: (Optional) A string representing a cell or formula to divide the sum by (e.g., "D1", "2").
   * @returns {string} The complete SUMIFS formula string.
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype.buildMonthlySumFormula = function(params) {
    const { sumRange, criteria, sharedDivisor } = params;
    
    const criteriaStrings = criteria.map(criterion => {
      if (criterion.type === 'date') {
        return this._buildDateCriteriaString(criterion.range, criterion.operator || '=', criterion.value);
      } else if (criterion.operator) {
        return this._buildCriteriaOperatorString(criterion.range, criterion.operator, criterion.value);
      }
      return this._buildCriteriaString(criterion.range, criterion.value);
    });
    
    const sumifs = `SUMIFS(${sumRange},${criteriaStrings.join(',')})`;
    
    if (sharedDivisor) {
      return `(${sumifs})/${sharedDivisor}`;
    }
    
    return sumifs;
  };

  /**
   * Builds a SUM formula for totaling values across a row.
   * @param {string} startCol - The starting column letter (e.g., "C").
   * @param {string} endCol - The ending column letter (e.g., "N").
   * @param {number} row - The row number.
   * @returns {string} The SUM formula string (e.g., "=SUM(C5:N5)").
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype.buildRowTotalFormula = function(startCol, endCol, row) {
    return `=SUM(${startCol}${row}:${endCol}${row})`;
  };

  /**
   * Builds an AVERAGE formula for averaging values across a row.
   * @param {string} startCol - The starting column letter (e.g., "C").
   * @param {string} endCol - The ending column letter (e.g., "N").
   * @param {number} row - The row number.
   * @returns {string} The AVERAGE formula string (e.g., "=AVERAGE(C5:N5)").
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype.buildRowAverageFormula = function(startCol, endCol, row) {
    return `=AVERAGE(${startCol}${row}:${endCol}${row})`;
  };

  /**
   * Builds a formula for net calculations (e.g., Income - Expenses).
   * @param {Array<{operation: 'add'|'subtract', reference: string}>} components - An array of objects,
   *   each specifying an operation ('add' or 'subtract') and a cell reference or value string.
   *   The first component's operation is ignored (assumed positive).
   * @returns {string} The net calculation formula string (e.g., "=A1-B1+C1").
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype.buildNetFormula = function(components) {
    const parts = components.map((comp, index) => {
      const prefix = index === 0 ? '' : (comp.operation === 'add' ? '+' : '-');
      return `${prefix}${comp.reference}`;
    });
    return `=${parts.join('')}`;
  };

  /**
   * Builds a formula reference to a cell, potentially on another sheet.
   * Handles sheet names with spaces or special characters by quoting them.
   * @param {string} sheet - The name of the sheet.
   * @param {string} column - The column letter (e.g., "A").
   * @param {number} row - The row number.
   * @returns {string} The complete cell reference string (e.g., "'Sheet Name'!A1").
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype.buildCellReference = function(sheet, column, row) {
    const sheetName = sheet.replace(/'/g, "''");
    return `'${sheetName}'!${column}${row}`;
  };

  /**
   * Builds an IF formula for conditional calculations.
   * @param {string} condition - The condition to evaluate (e.g., "A1>10").
   * @param {string} trueValue - The value or formula if the condition is true.
   * @param {string} falseValue - The value or formula if the condition is false.
   * @returns {string} The complete IF formula string (e.g., "IF(A1>10, B1, C1)").
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype.buildIfFormula = function(condition, trueValue, falseValue) {
    return `IF(${condition}, ${trueValue}, ${falseValue})`;
  };

  /**
   * Builds a percentage formula (numerator / denominator), wrapped in IFERROR to return 0 on division by zero.
   * @param {string} numerator - The cell reference or value for the numerator.
   * @param {string} denominator - The cell reference or value for the denominator.
   * @param {boolean} [absolute=false] - If true, uses ABS() for the numerator.
   * @returns {string} The percentage formula string (e.g., "=IFERROR(A1/B1,0)").
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype.buildPercentageFormula = function(numerator, denominator, absolute = false) {
    const num = absolute ? `ABS(${numerator})` : numerator;
    return `=IFERROR(${num}/${denominator},0)`;
  };

  /**
   * Builds a complex SUMIFS formula specifically for calculating category totals for a given month.
   * Used in financial overview sheets.
   * @param {{
   *   transactionSheet: string,
   *   amountColumn: string,
   *   typeColumn: string,
   *   categoryColumn: string,
   *   subcategoryColumn: string,
   *   dateColumn: string,
   *   typeValue: string,
   *   categoryValue?: string,
   *   subcategoryValue?: string,
   *   monthDate: Date,
   *   overviewSheetName: string,
   *   currentRow: number,
   *   showSubCategories: boolean
   * }} params - Parameters for building the category total formula.
   *   - `transactionSheet`: Name of the sheet containing transactions.
   *   - `amountColumn`: Column letter for amounts in transaction sheet.
   *   - `typeColumn`: Column letter for transaction types.
   *   - `categoryColumn`: Column letter for categories.
   *   - `subcategoryColumn`: Column letter for subcategories.
   *   - `dateColumn`: Column letter for dates.
   *   - `typeValue`: Cell reference (e.g., "$A5") or direct value for the type criteria.
   *   - `categoryValue`: (Optional) Cell reference or value for category.
   *   - `subcategoryValue`: (Optional) Cell reference or value for subcategory.
   *   - `monthDate`: A Date object representing any day in the target month.
   *   - `overviewSheetName`: Name of the sheet where this formula is placed (used for cell refs).
   *   - `currentRow`: The current row number on the overview sheet.
   *   - `showSubCategories`: Boolean indicating if subcategory criteria should be included.
   * @returns {string} The complete SUMIFS formula string for category totals.
   * @memberof FormulaBuilderModule
   */
  FormulaBuilderModuleConstructor.prototype.buildCategoryTotalFormula = function(params) {
    const {
      transactionSheet,
      amountColumn,
      typeColumn,
      categoryColumn,
      subcategoryColumn,
      dateColumn,
      typeValue,
      categoryValue,
      subcategoryValue,
      monthDate,
      overviewSheetName,
      currentRow,
      showSubCategories
    } = params;

    const month = monthDate.getMonth() + 1;
    const year = monthDate.getFullYear();
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);
    
    const criteria = [
      { range: `${transactionSheet}!${typeColumn}:${typeColumn}`, 
        value: `${overviewSheetName}!$A${currentRow}` },
      { range: `${transactionSheet}!${dateColumn}:${dateColumn}`, 
        operator: '>=', 
        value: Utilities.formatDate(startDate, Session.getScriptTimeZone(), this.config.getSection('LOCALE').DATE_FORMAT), 
        type: 'date' },
      { range: `${transactionSheet}!${dateColumn}:${dateColumn}`, 
        operator: '<=', 
        value: Utilities.formatDate(endDate, Session.getScriptTimeZone(), this.config.getSection('LOCALE').DATE_FORMAT), 
        type: 'date' }
    ];

    if (categoryValue) {
      criteria.push({ 
        range: `${transactionSheet}!${categoryColumn}:${categoryColumn}`, 
        value: `${overviewSheetName}!$B${currentRow}` 
      });
      
      if (subcategoryValue && showSubCategories) {
        criteria.push({ 
          range: `${transactionSheet}!${subcategoryColumn}:${subcategoryColumn}`, 
          value: `${overviewSheetName}!$C${currentRow}` 
        });
      }
    }

    return this.buildMonthlySumFormula({
      sumRange: `${transactionSheet}!${amountColumn}:${amountColumn}`,
      criteria: criteria
    });
  };

  return FormulaBuilderModuleConstructor;
})();
