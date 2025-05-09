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
      WANTS: 0.2,         // 30% for wants/pleasure
      EXTRA: 0.1,         // 20% for extras
      SAVINGS: 0.2,       // 20% for savings
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
        TYPE: 120,
        CATEGORY: 120,
        SUBCATEGORY: 150,
        SHARED: 60,
        MONTH: 70,
        AVERAGE: 80,
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
        // METRICS_BG: "#FFEBEE",      // Very light red for metrics section
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
      CURRENCY_LOCALE: "0", // Euro
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
  /**
   * Deeply merges properties from the source object into the target object.
   * @param {object} target The target object to merge into.
   * @param {object} source The source object to merge from.
   * @return {object} The modified target object.
   * @private
   */
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
     * Gets a deep copy of the complete configuration object, merging default and user-specific settings.
     * @return {object} The complete configuration object.
     *
     * @example
     * const currentConfig = FinancialPlanner.Config.get();
     * console.log(currentConfig.SHEETS.OVERVIEW); // Outputs: "Overview"
     */
    get: function() {
      // Return a deep copy of the merged configuration
      return JSON.parse(JSON.stringify(mergeConfig({}, DEFAULT_CONFIG)));
    },
    
    /**
     * Gets a specific top-level section from the configuration.
     * @param {string} section The name of the configuration section (e.g., 'SHEETS', 'COLORS').
     * @return {object} The requested configuration section, or an empty object if the section doesn't exist.
     *
     * @example
     * const sheetSettings = FinancialPlanner.Config.getSection('SHEETS');
     * console.log(sheetSettings.TRANSACTIONS); // Outputs: "Transactions"
     *
     * const nonExistent = FinancialPlanner.Config.getSection('NON_EXISTENT');
     * console.log(nonExistent); // Outputs: {}
     */
    getSection: function(section) {
      const config = this.get();
      return config[section] || {};
    },
    
    /**
     * Gets the sheet names configuration.
     * @return {object} An object containing sheet name configurations.
     *
     * @example
     * const sheetNames = FinancialPlanner.Config.getSheetNames();
     * console.log(sheetNames.OVERVIEW); // Outputs: "Overview"
     */
    getSheetNames: function() {
      return this.getSection('SHEETS');
    },
    
    /**
     * Gets the transaction types configuration.
     * @return {object} An object containing transaction type configurations.
     *
     * @example
     * const transactionTypes = FinancialPlanner.Config.getTransactionTypes();
     * console.log(transactionTypes.INCOME); // Outputs: "Income"
     */
    getTransactionTypes: function() {
      return this.getSection('TRANSACTION_TYPES');
    },
    
    /**
     * Gets the target rates configuration for expense categories.
     * @return {object} An object containing target rate configurations.
     *
     * @example
     * const targetRates = FinancialPlanner.Config.getTargetRates();
     * console.log(targetRates.ESSENTIALS); // Outputs: 0.5
     */
    getTargetRates: function() {
      return this.getSection('TARGET_RATES');
    },
    
    /**
     * Gets the UI elements configuration.
     * @return {object} An object containing UI configurations.
     *
     * @example
     * const uiConfig = FinancialPlanner.Config.getUI();
     * console.log(uiConfig.SUBCATEGORY_TOGGLE.LABEL_TEXT); // Outputs: "Show Sub-Categories"
     */
    getUI: function() {
      return this.getSection('UI');
    },
    
    /**
     * Gets the color configurations for UI elements and charts.
     * @return {object} An object containing color configurations.
     *
     * @example
     * const colors = FinancialPlanner.Config.getColors();
     * console.log(colors.TYPE_HEADERS.INCOME.BG); // Outputs: "#2E7D32"
     */
    getColors: function() {
      return this.getSection('COLORS');
    },
    
    /**
     * Gets the locale-specific configurations (e.g., currency symbol, date format).
     * @return {object} An object containing locale configurations.
     *
     * @example
     * const localeSettings = FinancialPlanner.Config.getLocale();
     * console.log(localeSettings.CURRENCY_SYMBOL); // Outputs: "€"
     */
    getLocale: function() {
      return this.getSection('LOCALE');
    },
    
    /**
     * Updates the in-memory user configuration by merging it with the provided settings.
     * Note: This updates the configuration for subsequent `get()` calls within the current session.
     * It does not persist changes.
     * @param {object} config The user-specific configuration object to merge.
     *
     * @example
     * FinancialPlanner.Config.update({
     *   SHEETS: { OVERVIEW: "My Custom Overview" },
     *   LOCALE: { CURRENCY_SYMBOL: "$" }
     * });
     * console.log(FinancialPlanner.Config.get().SHEETS.OVERVIEW); // Outputs: "My Custom Overview"
     * console.log(FinancialPlanner.Config.get().LOCALE.CURRENCY_SYMBOL); // Outputs: "$"
     */
    update: function(config) {
      userConfig = mergeConfig(userConfig, config);
    },
    
    /**
     * Resets any user-specific configuration overrides, reverting to the default configuration.
     *
     * @example
     * FinancialPlanner.Config.update({ SHEETS: { OVERVIEW: "Temporary Name" } });
     * // ... some operations ...
     * FinancialPlanner.Config.reset();
     * console.log(FinancialPlanner.Config.get().SHEETS.OVERVIEW); // Outputs: "Overview" (the default)
     */
    reset: function() {
      userConfig = {};
    }
  };
})();
