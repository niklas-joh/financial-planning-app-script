/**
 * @fileoverview Metrics Calculator Service for Financial Planning Tools.
 * Centralizes financial metric calculations.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const MetricsCalculatorModule = (function() {
  /**
   * Constructor for the MetricsCalculatorModule.
   * @param {object} configInstance - An instance of ConfigModule.
   * @constructor
   */
  function MetricsCalculatorModuleConstructor(configInstance) {
    this.config = configInstance;
  }

  /**
   * Calculates savings rate
   * @param {number} income Total income
   * @param {number} savings Total savings
   * @returns {number} Savings rate as decimal
   */
  MetricsCalculatorModuleConstructor.prototype.calculateSavingsRate = function(income, savings) {
    if (income === 0) return 0;
    return savings / income;
  };

  /**
   * Calculates expense rate by category
   * @param {number} expense Expense amount (can be negative)
   * @param {number} income Total income
   * @returns {number} Expense rate as decimal
   */
  MetricsCalculatorModuleConstructor.prototype.calculateExpenseRate = function(expense, income) {
    if (income === 0) return 0;
    return Math.abs(expense) / income;
  };

  /**
   * Calculates variance from target
   * @param {number} actual Actual value
   * @param {number} target Target value
   * @returns {number} Variance (actual - target)
   */
  MetricsCalculatorModuleConstructor.prototype.calculateVariance = function(actual, target) {
    return actual - target;
  };

  /**
   * Calculates percentage change
   * @param {number} current Current value
   * @param {number} previous Previous value
   * @returns {number} Percentage change as decimal
   */
  MetricsCalculatorModuleConstructor.prototype.calculatePercentageChange = function(current, previous) {
    if (previous === 0) return current > 0 ? 1 : 0;
    return (current - previous) / previous;
  };

  /**
   * Calculates net income after expenses
   * @param {number} income Total income
   * @param {number} expenses Total expenses (typically negative)
   * @returns {number} Net income
   */
  MetricsCalculatorModuleConstructor.prototype.calculateNetIncome = function(income, expenses) {
    return income + expenses; // expenses are negative
  };

  /**
   * Calculates allocatable income
   * @param {number} income Total income
   * @param {number} essentials Essential expenses (negative)
   * @param {number} wants Wants expenses (negative)
   * @param {number} savings Savings amount (positive)
   * @returns {number} Allocatable income
   */
  MetricsCalculatorModuleConstructor.prototype.calculateAllocatableIncome = function(income, essentials, wants, savings) {
    return income + essentials + wants - savings;
  };

  /**
   * Aggregates metrics by category
   * @param {Array} transactions Array of transaction objects
   * @param {string} categoryField Field name for category grouping
   * @returns {Object} Aggregated data by category
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
   * Calculates trend indicators
   * @param {number} current Current value
   * @param {number} previous Previous value
   * @param {Object} [thresholds] Threshold values for trend determination
   * @returns {Object} Trend information
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
   * Validates financial metrics
   * @param {Object} metrics Object containing financial metrics
   * @returns {Object} Validation result
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
   * Formats currency value
   * @param {number} value Value to format
   * @param {Object} [options] Formatting options
   * @returns {string} Formatted currency string
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
   * Calculates running totals
   * @param {Array} values Array of numeric values
   * @returns {Array} Array of running totals
   */
  MetricsCalculatorModuleConstructor.prototype.calculateRunningTotals = function(values) {
    let runningTotal = 0;
    return values.map(value => {
      runningTotal += value;
      return runningTotal;
    });
  };

  /**
   * Calculates moving average
   * @param {Array} values Array of numeric values
   * @param {number} period Period for moving average
   * @returns {Array} Array of moving averages
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