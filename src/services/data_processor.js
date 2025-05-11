/**
 * @fileoverview Data Processor Service for Financial Planning Tools.
 * Provides functionalities for extracting, transforming, filtering, and aggregating
 * financial transaction data. It includes an inner `DataProcessor` class that
 * operates on a given dataset.
 * This module is designed to be instantiated by `00_module_loader.js`.
 * @module services/data-processor
 */

/**
 * IIFE to encapsulate the DataProcessorModule logic.
 * @returns {function} The DataProcessorModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const DataProcessorModule = (function() {
  /**
   * Constructor for the DataProcessorModule.
   * This module serves as a factory for creating `DataProcessor` instances.
   * @param {ConfigModule} configInstance - An instance of ConfigModule.
   * @param {ErrorServiceModule} errorServiceInstance - An instance of ErrorServiceModule.
   * @constructor
   * @alias DataProcessorModule
   * @memberof module:services/data-processor
   */
  function DataProcessorModuleConstructor(configInstance, errorServiceInstance) {
    /**
     * Instance of ConfigModule.
     * @type {ConfigModule}
     * @private
     */
    this.config = configInstance;
    /**
     * Instance of ErrorServiceModule.
     * @type {ErrorServiceModule}
     * @private
     */
    this.errorService = errorServiceInstance;
  }

  /**
   * @classdesc An internal class responsible for performing various data processing
   * operations on a 2D array of financial data. Instances are created via
   * `DataProcessorModule.create()`.
   * @class DataProcessor
   * @private
   */
  class DataProcessor {
    /**
     * Creates an instance of DataProcessor.
     * @param {Array<Array<*>>} data - The raw 2D array of data, where the first row is headers.
     * @param {object} columnIndices - An object mapping standard column names
     *   (e.g., 'type', 'category', 'date', 'amount') to their respective
     *   zero-based column index in the `data` array.
     * @param {ConfigModule} config - An instance of the ConfigModule.
     * @param {ErrorServiceModule} errorService - An instance of the ErrorServiceModule.
     */
    constructor(data, columnIndices, config, errorService) {
      /**
       * The raw 2D array of data, including headers as the first row.
       * @type {Array<Array<*>>}
       */
      this.data = data;
      /**
       * An object mapping column names to their zero-based index.
       * @type {object}
       * @example { type: 0, category: 1, date: 3, amount: 4 }
       */
      this.indices = columnIndices;
      /**
       * Instance of ConfigModule.
       * @type {ConfigModule}
       * @private
       */
      this.config = config;
      /**
       * Instance of ErrorServiceModule.
       * @type {ErrorServiceModule}
       * @private
       */
      this.errorService = errorService;
      /**
       * The header row of the data.
       * @type {Array<string>}
       */
      this.headers = data[0];
    }

    /**
     * Extracts unique combinations of Type, Category, and optionally Sub-Category
     * from the transaction data.
     * @param {boolean} showSubCategories - If true, subcategories are included in the
     *   uniqueness check and the result. If false, subcategory is ignored and returned as empty.
     * @returns {Array<{type: string, category: string, subcategory: string}>}
     *   An array of objects, each representing a unique combination.
     */
    getUniqueCombinations(showSubCategories) {
      const seen = new Set();
      const combinations = [];
      
      for (let i = 1; i < this.data.length; i++) {
        const row = this.data[i];
        const type = row[this.indices.type];
        const category = row[this.indices.category];
        const subcategory = showSubCategories ? row[this.indices.subcategory] : '';
        
        if (!type || !category) continue;
        
        const key = `${type}|${category}|${subcategory}`;
        if (!seen.has(key)) {
          seen.add(key);
          combinations.push({
            type,
            category,
            subcategory: subcategory || ''
          });
        }
      }
      
      return combinations;
    }

    /**
     * Groups an array of category combinations (typically from `getUniqueCombinations`)
     * by their 'type' property. Also sorts combinations within each type.
     * @param {Array<{type: string, category: string, subcategory: string}>} combinations -
     *   An array of combination objects.
     * @returns {Object<string, Array<{type: string, category: string, subcategory: string}>>}
     *   An object where keys are types (e.g., "Income", "Expense") and values are
     *   arrays of sorted combination objects belonging to that type.
     */
    groupByType(combinations) {
      const grouped = {};
      
      combinations.forEach(combo => {
        if (!grouped[combo.type]) {
          grouped[combo.type] = [];
        }
        grouped[combo.type].push(combo);
      });
      
      // Sort within each group
      Object.keys(grouped).forEach(type => {
        grouped[type].sort((a, b) => {
          const catCompare = a.category.localeCompare(b.category);
          if (catCompare !== 0) return catCompare;
          return (a.subcategory || '').localeCompare(b.subcategory || '');
        });
      });
      
      return grouped;
    }

    /**
     * Filters the transaction data to include only rows within a specified date range.
     * The header row is excluded from the filtered result.
     * @param {Date} startDate - The start date for the filter (inclusive).
     * @param {Date} endDate - The end date for the filter (inclusive).
     * @returns {Array<Array<*>>} A new 2D array containing data rows (no header)
     *   that fall within the specified date range.
     */
    filterByDateRange(startDate, endDate) {
      return this.data.filter((row, index) => {
        if (index === 0) return false; // Skip header
        const date = new Date(row[this.indices.date]);
        return date >= startDate && date <= endDate;
      });
    }

    /**
     * Filters transaction data by a specific transaction type.
     * Excludes the header row.
     * @param {string} type - The transaction type to filter by (e.g., "Income", "Expense").
     * @returns {Array<Array<*>>} A new 2D array containing data rows of the specified type.
     */
    filterByType(type) {
      return this.data.filter((row, index) => {
        if (index === 0) return false;
        return row[this.indices.type] === type;
      });
    }

    /**
     * Filters transaction data by a specific category.
     * Excludes the header row.
     * @param {string} category - The category to filter by.
     * @returns {Array<Array<*>>} A new 2D array containing data rows of the specified category.
     */
    filterByCategory(category) {
      return this.data.filter((row, index) => {
        if (index === 0) return false;
        return row[this.indices.category] === category;
      });
    }

    /**
     * Filters transaction data based on a set of criteria.
     * Each key in the criteria object corresponds to a column name (e.g., 'type', 'category'),
     * and its value is the desired value for that column.
     * Excludes the header row.
     * @param {object} criteria - An object where keys are column names (must exist in `this.indices`)
     *   and values are the values to filter by.
     *   Example: `{ type: "Expense", category: "Food" }`
     * @returns {Array<Array<*>>} A new 2D array containing data rows matching all criteria.
     */
    filterByCriteria(criteria) {
      return this.data.filter((row, index) => {
        if (index === 0) return false;
        
        for (const [key, value] of Object.entries(criteria)) {
          if (this.indices[key] !== undefined && row[this.indices[key]] !== value) {
            return false;
          }
        }
        return true;
      });
    }

    /**
     * Aggregates transaction data by month, calculating total income, expenses,
     * savings, and collecting transactions for each month.
     * @returns {Object<string, {income: number, expenses: number, savings: number, transactions: Array<Array<*>>}>}
     *   An object where keys are month strings (e.g., "YYYY-M") and values are objects
     *   containing aggregated financial data for that month.
     */
    aggregateByMonth() {
      const monthlyData = {};
      
      for (let i = 1; i < this.data.length; i++) {
        const row = this.data[i];
        const date = new Date(row[this.indices.date]);
        const monthKey = `${date.getFullYear()}-${date.getMonth() + 1}`;
        
        if (!monthlyData[monthKey]) {
          monthlyData[monthKey] = {
            income: 0,
            expenses: 0,
            savings: 0,
            transactions: []
          };
        }
        
        const amount = parseFloat(row[this.indices.amount]) || 0;
        const type = row[this.indices.type];
        
        if (type === 'Income') {
          monthlyData[monthKey].income += amount;
        } else if (type === 'Savings') {
          monthlyData[monthKey].savings += Math.abs(amount);
        } else {
          monthlyData[monthKey].expenses += amount;
        }
        
        monthlyData[monthKey].transactions.push(row);
      }
      
      return monthlyData;
    }

    /**
     * Aggregates transaction data by category, calculating total amount, transaction count,
     * and collecting transactions for each category.
     * @returns {Object<string, {total: number, count: number, transactions: Array<Array<*>>}>}
     *   An object where keys are category names and values are objects containing
     *   aggregated data for that category.
     */
    aggregateByCategory() {
      const categoryData = {};
      
      for (let i = 1; i < this.data.length; i++) {
        const row = this.data[i];
        const category = row[this.indices.category];
        
        if (!category) continue;
        
        if (!categoryData[category]) {
          categoryData[category] = {
            total: 0,
            count: 0,
            transactions: []
          };
        }
        
        const amount = parseFloat(row[this.indices.amount]) || 0;
        categoryData[category].total += amount;
        categoryData[category].count++;
        categoryData[category].transactions.push(row);
      }
      
      return categoryData;
    }

    /**
     * Calculates monthly total amounts for a specific combination of type, category,
     * and subcategory over a hardcoded year (2024 - needs configuration).
     * @param {string} type - The transaction type (e.g., "Expense").
     * @param {string} [category=null] - Optional. The specific category.
     * @param {string} [subcategory=null] - Optional. The specific subcategory.
     * @returns {Object<number, number>} An object where keys are month indices (0-11)
     *   and values are the total amounts for that month and criteria.
     *   Note: The year is currently hardcoded to 2024.
     */
    getMonthlyTotals(type, category = null, subcategory = null) {
      const monthlyTotals = {};
      
      for (let month = 0; month < 12; month++) {
        const year = 2024; // This should be configurable
        const startDate = new Date(year, month, 1);
        const endDate = new Date(year, month + 1, 0);
        
        const transactions = this.filterByDateRange(startDate, endDate)
          .filter((row, index) => {
            if (index === 0) return false;
            if (row[this.indices.type] !== type) return false;
            if (category && row[this.indices.category] !== category) return false;
            if (subcategory && row[this.indices.subcategory] !== subcategory) return false;
            return true;
          });
        
        const total = transactions.reduce((sum, row) => {
          return sum + (parseFloat(row[this.indices.amount]) || 0);
        }, 0);
        
        monthlyTotals[month] = total;
      }
      
      return monthlyTotals;
    }

    /**
     * Validates that the necessary columns (as defined by `requiredColumns`)
     * are present in the provided data's column indices.
     * @returns {boolean} True if the structure is valid.
     * @throws {Error} An error (created by ErrorService) if required columns are missing.
     *   The error object includes details like the missing columns and current headers.
     */
    validateStructure() {
      const requiredColumns = ['type', 'category', 'date', 'amount'];
      const missing = requiredColumns.filter(col => this.indices[col] === -1);
      
      if (missing.length > 0) {
        throw this.errorService.create(
          `Required columns not found: ${missing.join(', ')}`,
          { headers: this.headers, severity: 'high' }
        );
      }
      
      return true;
    }

    /**
     * Determines the indices of predefined columns (Type, Category, Sub-Category, Date, Amount, Shared)
     * based on an array of header strings.
     * @param {Array<string>} headers - An array of strings representing the column headers.
     * @returns {{type: number, category: number, subcategory: number, date: number, amount: number, shared: number}}
     *   An object mapping standard column names to their found zero-based index.
     *   Returns -1 for columns not found.
     * @static
     */
    static getColumnIndices(headers) {
      const indices = {
        type: headers.indexOf("Type"),
        category: headers.indexOf("Category"),
        subcategory: headers.indexOf("Sub-Category"),
        date: headers.indexOf("Date"),
        amount: headers.indexOf("Amount"),
        shared: headers.indexOf("Shared")
      };
      
      return indices;
    }

    /**
     * Maps a data row (as an array) to an object with named properties,
     * based on the current column indices. Also parses the amount to a float.
     * @param {Array<*>} row - A single data row array.
     * @returns {{type: string, category: string, subcategory: string, date: *, amount: number, shared: *}}
     *   An object representation of the row.
     */
    mapRowToObject(row) {
      return {
        type: row[this.indices.type],
        category: row[this.indices.category],
        subcategory: row[this.indices.subcategory],
        date: row[this.indices.date],
        amount: parseFloat(row[this.indices.amount]) || 0,
        shared: row[this.indices.shared]
      };
    }

    /**
     * Calculates summary statistics for the entire dataset, including total income,
     * expenses, savings, transaction count, and the date range of transactions.
     * @returns {{totalIncome: number, totalExpenses: number, totalSavings: number, transactionCount: number, dateRange: {start: Date|null, end: Date|null}}}
     *   An object containing the summary statistics.
     */
    getSummaryStatistics() {
      const stats = {
        totalIncome: 0,
        totalExpenses: 0,
        totalSavings: 0,
        transactionCount: 0,
        dateRange: {
          start: null,
          end: null
        }
      };
      
      for (let i = 1; i < this.data.length; i++) {
        const row = this.data[i];
        const amount = parseFloat(row[this.indices.amount]) || 0;
        const type = row[this.indices.type];
        const date = new Date(row[this.indices.date]);
        
        if (type === 'Income') {
          stats.totalIncome += amount;
        } else if (type === 'Savings') {
          stats.totalSavings += Math.abs(amount);
        } else {
          stats.totalExpenses += amount;
        }
        
        stats.transactionCount++;
        
        if (!stats.dateRange.start || date < stats.dateRange.start) {
          stats.dateRange.start = date;
        }
        if (!stats.dateRange.end || date > stats.dateRange.end) {
          stats.dateRange.end = date;
        }
      }
      
      return stats;
    }
  }

  // Public API
  /**
   * Creates and returns a new instance of the internal `DataProcessor` class,
   * initialized with the given data and column indices.
   * @param {Array<Array<*>>} data - The raw 2D array of transaction data,
   *   where the first row contains headers.
   * @param {object} columnIndices - An object mapping standard column names
   *   (e.g., 'type', 'category') to their respective zero-based column indices.
   *   Typically generated by `DataProcessor.getColumnIndices()`.
   * @returns {DataProcessor} A new instance of the `DataProcessor` class.
   * @memberof DataProcessorModule
   */
  DataProcessorModuleConstructor.prototype.create = function(data, columnIndices) {
    return new DataProcessor(data, columnIndices, this.config, this.errorService);
  };

  /**
   * A utility function to get column indices from a header row.
   * This is a convenience method that delegates to the static method
   * on the internal `DataProcessor` class.
   * @param {Array<string>} headers - An array of strings representing the column headers.
   * @returns {{type: number, category: number, subcategory: number, date: number, amount: number, shared: number}}
   *   An object mapping standard column names to their found zero-based index.
   *   Returns -1 for columns not found.
   * @memberof DataProcessorModule
   */
  DataProcessorModuleConstructor.prototype.getColumnIndices = function(headers) {
    return DataProcessor.getColumnIndices(headers);
  };

  return DataProcessorModuleConstructor;
})();
