/**
 * Financial Planning Tools - Configuration Module
 * 
 * This file contains centralized configuration settings for the Financial Planning Tools project.
 * It helps reduce duplication and makes it easier to manage configuration across the application.
 */

// Create the Config module within the FinancialPlanner namespace
FinancialPlanner.Config = (function() {
  // Default configuration
  const DEFAULT_CONFIG = {
    SHEETS: {
      OVERVIEW: "Overview",
      TRANSACTIONS: "Transactions",
      DROPDOWNS: "Dropdowns",
      ERROR_LOG: "Error Log",
      ANALYSIS: "Analysis",
      SETTINGS: "Settings"
    },
    TRANSACTION_TYPES: {
      INCOME: "Income",
      ESSENTIALS: "Essentials",
      WANTS: "Wants/Pleasure",
      EXTRA: "Extra", 
      SAVINGS: "Savings"
    },
    // Transaction types in preferred display order
    TYPE_ORDER: ["Income", "Essentials", "Wants/Pleasure", "Extra", "Savings"],
    // Types considered as expenses (for shared expense calculations)
    EXPENSE_TYPES: ["Essentials", "Wants/Pleasure", "Extra"],
    // Default target rates for expense categories
    TARGET_RATES: {
      ESSENTIALS: 0.5,    // 50% for essentials
      WANTS: 0.3,         // 30% for wants/pleasure
      EXTRA: 0.2,         // 20% for extras
      DEFAULT: 0.2        // 20% default
    },
    // Header structure for the overview sheet
    HEADERS: [
      "Type", "Category", "Sub-Category", "Shared?", 
      "Jan-24", "Feb-24", "Mar-24", "Apr-24", 
      "May-24", "Jun-24", "Jul-24", "Aug-24", 
      "Sep-24", "Oct-24", "Nov-24", "Dec-24", "Total", "Average"
    ],
    // UI element positions and names
    UI: {
      SUBCATEGORY_TOGGLE: {
        LABEL_CELL: "S1",
        CHECKBOX_CELL: "T1",
        LABEL_TEXT: "Show Sub-Categories",
        NOTE_TEXT: "Toggle to show or hide sub-categories in the overview sheet"
      },
      COLUMN_WIDTHS: {
        TYPE: 150,
        CATEGORY: 150,
        SUBCATEGORY: 150,
        SHARED: 80,
        MONTH: 90,
        AVERAGE: 100,
        EXPENSE_CATEGORY: 150,
        AMOUNT: 100,
        RATE: 80
      }
    },
    COLORS: {
      // Type header colors
      TYPE_HEADERS: {
        INCOME: {
          BG: "#2E7D32",      // Green for Income
          FONT: "#FFFFFF"     // White text
        },
        ESSENTIALS: {
          BG: "#1976D2",      // Blue for Essentials
          FONT: "#FFFFFF"     // White text
        },
        WANTS_PLEASURE: {
          BG: "#FFA000",      // Amber for Wants/Pleasure
          FONT: "#FFFFFF"     // White text
        },
        EXTRA: {
          BG: "#7B1FA2",      // Purple for Extra
          FONT: "#FFFFFF"     // White text
        },
        SAVINGS: {
          BG: "#1565C0",      // Blue for Savings
          FONT: "#FFFFFF"     // White text
        },
        DEFAULT: {
          BG: "#424242",      // Dark gray for other types
          FONT: "#FFFFFF"     // White text
        }
      },
      // UI element colors
      UI: {
        HEADER_BG: "#C62828",       // Deep red for headers
        HEADER_FONT: "#FFFFFF",     // White text for headers
        METRICS_BG: "#FFEBEE",      // Very light red for metrics section
        BORDER: "#FF8F00",          // Amber for borders
        INCOME_FONT: "#388E3C",     // Green for income values
        EXPENSE_FONT: "#D32F2F",    // Red for expense values
        SAVINGS_FONT: "#1565C0",    // Blue for savings values
        NEUTRAL_FONT: "#424242",    // Dark gray for neutral values
        NET_BG: "#424242",          // Dark gray for net calculations
        NET_FONT: "#FFFFFF"         // White text for net calculations
      },
      // Chart colors
      CHART: {
        SERIES: [
          "#C62828", // Red (for Essentials)
          "#FF8F00", // Amber (for Wants/Pleasure)
          "#1565C0", // Blue (for Extra)
          "#2E7D32", // Green
          "#6A1B9A", // Purple
          "#E64A19", // Deep Orange
          "#00695C", // Teal
          "#5D4037"  // Brown
        ],
        TITLE: "#424242",
        TEXT: "#424242"
      }
    },
    // Cache configuration
    CACHE: {
      ENABLED: true,
      KEYS: {
        CATEGORY_COMBINATIONS: "finance_overview_categories",
        GROUPED_COMBINATIONS: "finance_overview_grouped"
      },
      EXPIRY_SECONDS: 21600 // 6 hours
    },
    // Locale settings
    LOCALE: {
      CURRENCY_SYMBOL: "€",
      CURRENCY_LOCALE: "2", // Euro
      DATE_FORMAT: "yyyy-MM-dd"
    },
    // Performance settings
    PERFORMANCE: {
      BATCH_SIZE: 50, // Number of rows to process in one batch
      USE_BATCH_OPERATIONS: true
    }
  };
  
  // Private variables
  let userConfig = {};
  
  // Private methods
  function mergeConfig(target, source) {
    for (const key in source) {
      if (source.hasOwnProperty(key)) {
        if (source[key] instanceof Object && !(source[key] instanceof Array) && target[key]) {
          // Recursively merge nested objects
          target[key] = mergeConfig(target[key], source[key]);
        } else {
          // Replace or add simple values
          target[key] = source[key];
        }
      }
    }
    return target;
  }
  
  // Public API
  return {
    /**
     * Gets the complete configuration object
     * @return {Object} The complete configuration
     */
    get: function() {
      // Return a deep copy of the merged configuration
      return JSON.parse(JSON.stringify(mergeConfig({}, DEFAULT_CONFIG)));
    },
    
    /**
     * Gets a specific configuration section
     * @param {String} section - The section name (e.g., 'SHEETS', 'COLORS')
     * @return {Object} The requested configuration section
     */
    getSection: function(section) {
      const config = this.get();
      return config[section] || {};
    },
    
    /**
     * Gets sheet names configuration
     * @return {Object} Sheet names configuration
     */
    getSheetNames: function() {
      return this.getSection('SHEETS');
    },
    
    /**
     * Gets transaction types configuration
     * @return {Object} Transaction types configuration
     */
    getTransactionTypes: function() {
      return this.getSection('TRANSACTION_TYPES');
    },
    
    /**
     * Gets target rates configuration
     * @return {Object} Target rates configuration
     */
    getTargetRates: function() {
      return this.getSection('TARGET_RATES');
    },
    
    /**
     * Gets UI configuration
     * @return {Object} UI configuration
     */
    getUI: function() {
      return this.getSection('UI');
    },
    
    /**
     * Gets colors configuration
     * @return {Object} Colors configuration
     */
    getColors: function() {
      return this.getSection('COLORS');
    },
    
    /**
     * Gets locale configuration
     * @return {Object} Locale configuration
     */
    getLocale: function() {
      return this.getSection('LOCALE');
    },
    
    /**
     * Updates the configuration with user-specific settings
     * @param {Object} config - The user configuration to merge
     */
    update: function(config) {
      userConfig = mergeConfig(userConfig, config);
    },
    
    /**
     * Resets the configuration to default values
     */
    reset: function() {
      userConfig = {};
    }
  };
})();
