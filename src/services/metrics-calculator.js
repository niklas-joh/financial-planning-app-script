/**
 * @fileoverview Metrics Calculator Service for Financial Planning Tools.
 * Centralizes common financial metric calculations such as savings rate,
 * expense rates, variance, percentage change, and more.
 * This module is designed to be instantiated by `00_module_loader.js`.
 * @module services/metrics-calculator
 */

/**
 * IIFE to encapsulate the MetricsCalculatorModule logic.
 * @returns {function} The MetricsCalculatorModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const MetricsCalculatorModule = (function() {
  /**
   * Constructor for the MetricsCalculatorModule.
   * @param {ConfigModule} configInstance - An instance of ConfigModule, used for locale-specific formatting.
   * @constructor
   * @alias MetricsCalculatorModule
   * @memberof module:services/metrics-calculator
   */
  function MetricsCalculatorModuleConstructor(configInstance) {
    /**
     * Instance of ConfigModule.
     * @type {ConfigModule}
     * @private
     */
    this.config = configInstance;
  }

  /**
   * Calculates the savings rate.
   * Savings rate is defined as (Total Savings / Total Income).
   * @param {number} income - Total income. Must be non-negative.
   * @param {number} savings - Total savings.
   * @returns {number} The savings rate as a decimal (e.g., 0.1 for 10%). Returns 0 if income is 0.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculateSavingsRate = function(income, savings) {
    if (income === 0) return 0;
    return savings / income;
  };

  /**
   * Calculates the expense rate for a given expense amount relative to total income.
   * Expense rate is defined as (Absolute Expense Amount / Total Income).
   * @param {number} expense - The expense amount. Can be positive or negative; its absolute value is used.
   * @param {number} income - Total income. Must be non-negative.
   * @returns {number} The expense rate as a decimal (e.g., 0.2 for 20%). Returns 0 if income is 0.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculateExpenseRate = function(expense, income) {
    if (income === 0) return 0;
    return Math.abs(expense) / income;
  };

  /**
   * Calculates the variance between an actual value and a target value.
   * Variance is defined as (Actual Value - Target Value).
   * @param {number} actual - The actual observed value.
   * @param {number} target - The target or budgeted value.
   * @returns {number} The difference between actual and target.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculateVariance = function(actual, target) {
    return actual - target;
  };

  /**
   * Calculates the percentage change between a current value and a previous value.
   * Percentage change is ((Current - Previous) / Previous).
   * @param {number} current - The current value.
   * @param {number} previous - The previous value.
   * @returns {number} The percentage change as a decimal (e.g., 0.05 for 5% increase).
   *   Returns 1 (100%) if previous is 0 and current is positive, 0 otherwise if previous is 0.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculatePercentageChange = function(current, previous) {
    if (previous === 0) return current > 0 ? 1 : 0;
    return (current - previous) / previous;
  };

  /**
   * Calculates net income (or loss).
   * Net income is (Total Income + Total Expenses). Assumes expenses are provided as negative values.
   * @param {number} income - Total income.
   * @param {number} expenses - Total expenses (conventionally a negative value).
   * @returns {number} The net income.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculateNetIncome = function(income, expenses) {
    return income + expenses; // expenses are negative
  };

  /**
   * Calculates allocatable income, often used in budgeting.
   * Allocatable income = Total Income + Essential Expenses + Wants Expenses - Savings.
   * Assumes expense parameters are negative and savings are positive.
   * @param {number} income - Total income.
   * @param {number} essentials - Total essential expenses (typically a negative value).
   * @param {number} wants - Total discretionary expenses (wants, typically a negative value).
   * @param {number} savings - Total amount allocated to savings (typically a positive value).
   * @returns {number} The remaining allocatable income.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculateAllocatableIncome = function(income, essentials, wants, savings) {
    return income + essentials + wants - savings;
  };

  /**
   * Aggregates transactions by a specified category field.
   * For each category, it calculates the count of transactions, total amount,
   * average amount, and collects all transactions belonging to that category.
   * @param {Array<object>} transactions - An array of transaction objects. Each object
   *   is expected to have an `amount` property (number) and a property matching `categoryField`.
   * @param {string} categoryField - The name of the property within each transaction object
   *   to use for grouping (e.g., "category", "type").
   * @returns {Object<string, {count: number, total: number, transactions: Array<object>, average: number}>}
   *   An object where keys are the unique values from `categoryField` and values are objects
   *   containing `count`, `total` amount, an array of `transactions`, and `average` amount for that category.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.aggregateByCategory = function(transactions, categoryField) {
    return transactions.reduce((acc, transaction) => {
      const category = transaction[categoryField];
      if (!acc[category]) {
        acc[category] = { 
          count: 0, 
          total: 0, 
          transactions: [],
          average: 0
        };
      }
      acc[category].count++;
      acc[category].total += transaction.amount;
      acc[category].transactions.push(transaction);
      acc[category].average = acc[category].total / acc[category].count;
      return acc;
    }, {});
  };

  /**
   * Calculates a trend indicator based on current and previous values and defined thresholds.
   * @param {number} current - The current period's value.
   * @param {number} previous - The previous period's value.
   * @param {{increase: number, decrease: number}} [thresholds={increase: 0.1, decrease: -0.1}] -
   *   Optional. An object defining the percentage change thresholds for 'up' and 'down' trends.
   *   `increase` should be positive (e.g., 0.1 for +10%), `decrease` should be negative (e.g., -0.1 for -10%).
   * @returns {{direction: 'up'|'down'|'stable', percentage: number, indicator: string, color: string}}
   *   An object describing the trend:
   *   - `direction`: 'up', 'down', or 'stable'.
   *   - `percentage`: The calculated percentage change (absolute for 'down' trend).
   *   - `indicator`: A visual indicator ('↑', '↓', '→').
   *   - `color`: A hex color code representing the trend (e.g., red for increase, green for decrease).
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculateTrend = function(current, previous, thresholds = { increase: 0.1, decrease: -0.1 }) {
    const change = this.calculatePercentageChange(current, previous);
    
    if (change > thresholds.increase) {
      return { 
        direction: 'up', 
        percentage: change, 
        indicator: '↑',
        color: '#CC0000' // Red for expense increase
      };
    } else if (change < thresholds.decrease) {
      return { 
        direction: 'down', 
        percentage: Math.abs(change), 
        indicator: '↓',
        color: '#006600' // Green for expense decrease
      };
    } else {
      return { 
        direction: 'stable', 
        percentage: change, 
        indicator: '→',
        color: '#666666' // Gray for stable
      };
    }
  };

  /**
   * Validates a set of financial metrics against common sense rules.
   * @param {object} metrics - An object containing financial metrics to validate. Expected properties might include:
   *   - `income`: Total income.
   *   - `savingsRate`: Calculated savings rate.
   *   - `expenseRates`: An object where keys are expense categories and values are their rates.
   * @returns {{valid: boolean, errors: Array<string>, warnings: Array<string>}}
   *   An object containing:
   *   - `valid`: True if no errors were found, false otherwise.
   *   - `errors`: An array of error message strings.
   *   - `warnings`: An array of warning message strings.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.validateMetrics = function(metrics) {
    const errors = [];
    const warnings = [];
    
    if (metrics.income < 0) {
      errors.push('Income cannot be negative');
    }
    
    if (metrics.savingsRate > 1) {
      errors.push('Savings rate cannot exceed 100%');
    }
    
    if (metrics.savingsRate < 0) {
      warnings.push('Negative savings rate indicates deficit');
    }
    
    const totalExpenseRate = Object.values(metrics.expenseRates || {})
      .reduce((sum, rate) => sum + rate, 0);
      
    if (totalExpenseRate > 1) {
      warnings.push('Total expense rate exceeds 100%');
    }
    
    return {
      valid: errors.length === 0,
      errors: errors,
      warnings: warnings
    };
  };

  /**
   * Formats a numeric value as a currency string based on locale settings from ConfigModule.
   * @param {number} value - The numeric value to format.
   * @param {{symbol?: string, decimals?: number}} [options={}] - Optional formatting options.
   *   - `symbol`: Override the default currency symbol from locale.
   *   - `decimals`: Override the default number of decimal places (default is 0).
   * @returns {string} The formatted currency string (e.g., "$1,234", "€-50.00").
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.formatCurrency = function(value, options = {}) {
    const locale = this.config.getLocale();
    const symbol = options.symbol || locale.CURRENCY_SYMBOL;
    const decimals = options.decimals !== undefined ? options.decimals : 0;
    
    const formatted = Math.abs(value).toFixed(decimals);
    const sign = value < 0 ? '-' : '';
    
    return `${sign}${symbol}${formatted}`;
  };

  /**
   * Calculates the running total for an array of numeric values.
   * @param {Array<number>} values - An array of numbers.
   * @returns {Array<number>} An array of the same length, where each element
   *   is the cumulative sum up to that point in the input array.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculateRunningTotals = function(values) {
    let runningTotal = 0;
    return values.map(value => {
      runningTotal += value;
      return runningTotal;
    });
  };

  /**
   * Calculates the moving average for an array of numeric values over a specified period.
   * @param {Array<number>} values - An array of numbers.
   * @param {number} period - The number of values to include in each average calculation (window size).
   * @returns {Array<number|null>} An array of moving averages. Elements before the first
   *   full period will be `null`.
   * @memberof MetricsCalculatorModule
   */
  MetricsCalculatorModuleConstructor.prototype.calculateMovingAverage = function(values, period) {
    const movingAverages = [];
    
    for (let i = 0; i < values.length; i++) {
      if (i < period - 1) {
        movingAverages.push(null);
      } else {
        const sum = values.slice(i - period + 1, i + 1).reduce((a, b) => a + b, 0);
        movingAverages.push(sum / period);
      }
    }
    
    return movingAverages;
  };

  return MetricsCalculatorModuleConstructor;
})();
