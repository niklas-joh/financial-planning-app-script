/**
 * @fileoverview Formula Builder Service for Financial Planning Tools.
 * Centralizes spreadsheet formula construction logic.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const FormulaBuilderModule = (function() {
  /**
   * Constructor for the FormulaBuilderModule.
   * @param {object} configInstance - An instance of ConfigModule.
   * @constructor
   */
  function FormulaBuilderModuleConstructor(configInstance) {
    this.config = configInstance;
  }

  // Private helper methods
  FormulaBuilderModuleConstructor.prototype._buildCriteriaString = function(criteriaRange, criteriaValue) {
    return `${criteriaRange},"${criteriaValue}"`;
  };

  FormulaBuilderModuleConstructor.prototype._buildDateCriteriaString = function(dateRange, operator, dateValue) {
    return `${dateRange},"${operator}${dateValue}"`;
  };

  FormulaBuilderModuleConstructor.prototype._buildCriteriaOperatorString = function(criteriaRange, operator, criteriaValue) {
    return `${criteriaRange},"${operator}${criteriaValue}"`;
  };

  // Public API methods

  /**
   * Builds a SUMIFS formula for monthly calculations
   * @param {Object} params Formula parameters
   * @param {string} params.sumRange Range to sum
   * @param {Array} params.criteria Array of criteria objects
   * @param {string} [params.sharedDivisor] Optional divisor formula
   * @returns {string} Complete formula string
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
   * Builds a SUM formula for row totals
   * @param {string} startCol Starting column letter
   * @param {string} endCol Ending column letter
   * @param {number} row Row number
   * @returns {string} Complete formula string
   */
  FormulaBuilderModuleConstructor.prototype.buildRowTotalFormula = function(startCol, endCol, row) {
    return `=SUM(${startCol}${row}:${endCol}${row})`;
  };

  /**
   * Builds an AVERAGE formula for row averages
   * @param {string} startCol Starting column letter
   * @param {string} endCol Ending column letter
   * @param {number} row Row number
   * @returns {string} Complete formula string
   */
  FormulaBuilderModuleConstructor.prototype.buildRowAverageFormula = function(startCol, endCol, row) {
    return `=AVERAGE(${startCol}${row}:${endCol}${row})`;
  };

  /**
   * Builds a formula for net calculations
   * @param {Array} components Array of component objects with operation and reference
   * @returns {string} Complete formula string
   */
  FormulaBuilderModuleConstructor.prototype.buildNetFormula = function(components) {
    const parts = components.map((comp, index) => {
      const prefix = index === 0 ? '' : (comp.operation === 'add' ? '+' : '-');
      return `${prefix}${comp.reference}`;
    });
    return `=${parts.join('')}`;
  };

  /**
   * Builds a formula reference to another cell
   * @param {string} sheet Sheet name
   * @param {string} column Column letter
   * @param {number} row Row number
   * @returns {string} Complete cell reference
   */
  FormulaBuilderModuleConstructor.prototype.buildCellReference = function(sheet, column, row) {
    const sheetName = sheet.replace(/'/g, "''");
    return `'${sheetName}'!${column}${row}`;
  };

  /**
   * Builds an IF formula for conditional calculations
   * @param {string} condition Condition to evaluate
   * @param {string} trueValue Value if true
   * @param {string} falseValue Value if false
   * @returns {string} Complete IF formula
   */
  FormulaBuilderModuleConstructor.prototype.buildIfFormula = function(condition, trueValue, falseValue) {
    return `IF(${condition}, ${trueValue}, ${falseValue})`;
  };

  /**
   * Builds a percentage formula
   * @param {string} numerator Numerator reference
   * @param {string} denominator Denominator reference
   * @param {boolean} [absolute=false] Whether to use absolute value for numerator
   * @returns {string} Complete formula string
   */
  FormulaBuilderModuleConstructor.prototype.buildPercentageFormula = function(numerator, denominator, absolute = false) {
    const num = absolute ? `ABS(${numerator})` : numerator;
    return `=IFERROR(${num}/${denominator},0)`;
  };

  /**
   * Builds a SUMIFS formula for category totals
   * @param {Object} params Formula parameters
   * @returns {string} Complete formula string
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