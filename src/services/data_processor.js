/**
 * @fileoverview Data Processor Service for Financial Planning Tools.
 * Handles data extraction and transformation.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const DataProcessorModule = (function() {
  /**
   * Constructor for the DataProcessorModule.
   * @param {object} configInstance - An instance of ConfigModule.
   * @param {object} errorServiceInstance - An instance of ErrorServiceModule.
   * @constructor
   */
  function DataProcessorModuleConstructor(configInstance, errorServiceInstance) {
    this.config = configInstance;
    this.errorService = errorServiceInstance;
  }

  /**
   * Internal DataProcessor class
   */
  class DataProcessor {
    constructor(data, columnIndices, config, errorService) {
      this.data = data;
      this.indices = columnIndices;
      this.config = config;
      this.errorService = errorService;
      this.headers = data[0];
    }

    /**
     * Extracts unique category combinations
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
     * Groups combinations by type
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
     * Filters transactions by date range
     */
    filterByDateRange(startDate, endDate) {
      return this.data.filter((row, index) => {
        if (index === 0) return false; // Skip header
        const date = new Date(row[this.indices.date]);
        return date >= startDate && date <= endDate;
      });
    }

    /**
     * Filters transactions by type
     */
    filterByType(type) {
      return this.data.filter((row, index) => {
        if (index === 0) return false;
        return row[this.indices.type] === type;
      });
    }

    /**
     * Filters transactions by category
     */
    filterByCategory(category) {
      return this.data.filter((row, index) => {
        if (index === 0) return false;
        return row[this.indices.category] === category;
      });
    }

    /**
     * Filters transactions by multiple criteria
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
     * Aggregates data by month
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
     * Aggregates by category
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
     * Gets monthly totals for a specific type/category combination
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
     * Validates data structure
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
     * Gets column indices from headers
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
     * Maps row data to object
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
     * Gets summary statistics
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
  DataProcessorModuleConstructor.prototype.create = function(data, columnIndices) {
    return new DataProcessor(data, columnIndices, this.config, this.errorService);
  };

  DataProcessorModuleConstructor.prototype.getColumnIndices = function(headers) {
    return DataProcessor.getColumnIndices(headers);
  };

  return DataProcessorModuleConstructor;
})();