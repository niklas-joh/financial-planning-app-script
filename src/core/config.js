/**
 * @fileoverview Configuration Module for Financial Planning Tools.
 * Provides centralized configuration settings for the application.
 * @module core/config
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Configuration Service - Provides centralized configuration management.
 * Uses IIFE to keep DEFAULT_CONFIG private via closure.
 * @namespace FinancialPlanner.Config
 */
FinancialPlanner.Config = (function() {
  /**
   * @const {object} DEFAULT_CONFIG
   * @private
   * @description Default configuration settings for the application.
   * This object serves as the base configuration and can be overridden by user-specific settings.
   * It includes sheet names, transaction types, UI settings, color schemes, cache settings,
   * locale information, and performance parameters.
   */
  const DEFAULT_CONFIG = {
    /** 
     * @property {object} SHEETS Defines the names of various sheets used in the spreadsheet.
     * Keys are internal identifiers, values are display names.
     */
    SHEETS: {
      OVERVIEW: 'Overview',
      TRANSACTIONS: 'Transactions',
      DROPDOWNS: 'Dropdowns',
      ERROR_LOG: 'Error Log',
      ANALYSIS: 'Analysis',
      SETTINGS: 'Settings',
    },
    /** 
     * @property {object} TRANSACTION_TYPES Defines categories for financial transactions.
     * Keys are internal identifiers, values are display names.
     */
    TRANSACTION_TYPES: {
      INCOME: 'Income',
      ESSENTIALS: 'Essentials',
      WANTS: 'Wants/Pleasure',
      EXTRA: 'Extra',
      SAVINGS: 'Savings',
    },
    /** @property {string[]} TYPE_ORDER Specifies the display order for transaction types. */
    TYPE_ORDER: ['Income', 'Essentials', 'Wants/Pleasure', 'Extra', 'Savings'],
    /** @property {string[]} EXPENSE_TYPES Lists transaction types considered as expenses. */
    EXPENSE_TYPES: ['Essentials', 'Wants/Pleasure', 'Extra'],
    /** 
     * @property {object} TARGET_RATES Defines target allocation rates (as decimals, e.g., 0.5 for 50%) 
     * for different expense/savings categories.
     */
    TARGET_RATES: {
      ESSENTIALS: 0.5,
      WANTS: 0.2,
      EXTRA: 0.1,
      SAVINGS: 0.2,
      DEFAULT: 0.2, // Default rate if a specific category is not listed
    },
    /** @property {string[]} HEADERS Defines the column headers for transaction and overview sheets. */
    HEADERS: [
      'Type', 'Category', 'Sub-Category', 'Shared?',
      'Jan-24', 'Feb-24', 'Mar-24', 'Apr-24',
      'May-24', 'Jun-24', 'Jul-24', 'Aug-24',
      'Sep-24', 'Oct-24', 'Nov-24', 'Dec-24', 'Total', 'Average',
    ],
    /** @property {object} UI Contains settings related to the user interface. */
    UI: {
      /** @property {object} SUBCATEGORY_TOGGLE Settings for the sub-category visibility toggle. */
      SUBCATEGORY_TOGGLE: {
        /** @property {string} LABEL_CELL Cell for the toggle label (e.g., 'S1'). */
        LABEL_CELL: 'S1',
        /** @property {string} CHECKBOX_CELL Cell for the toggle checkbox (e.g., 'T1'). */
        CHECKBOX_CELL: 'T1',
        /** @property {string} LABEL_TEXT Display text for the toggle label. */
        LABEL_TEXT: 'Show Sub-Categories',
        /** @property {string} NOTE_TEXT Explanatory note for the toggle. */
        NOTE_TEXT: 'Toggle to show or hide sub-categories in the overview sheet',
      },
      /** 
       * @property {object} COLUMN_WIDTHS Defines default column widths (in pixels) for various columns in sheets.
       * The keys represent the column identifier and values are the width.
       */
      COLUMN_WIDTHS: {
        TYPE: 120,
        CATEGORY: 120,
        SUBCATEGORY: 150,
        SHARED: 60,
        MONTH: 60,
        AVERAGE: 80,
        EXPENSE_CATEGORY: 150,
        AMOUNT: 100,
        RATE: 80,
      },
    },
    /** @property {object} COLORS Defines color schemes for UI elements and charts. */
    COLORS: {
      /** 
       * @property {object} TYPE_HEADERS Colors for headers based on transaction type. 
       * Each key (e.g., INCOME) has BG (background) and FONT (font color) properties.
       */
      TYPE_HEADERS: {
        INCOME: { BG: '#2E7D32', FONT: '#FFFFFF' }, // Dark Green BG, White Font
        ESSENTIALS: { BG: '#1976D2', FONT: '#FFFFFF' }, // Dark Blue BG, White Font
        WANTS_PLEASURE: { BG: '#FFA000', FONT: '#FFFFFF' }, // Amber BG, White Font
        EXTRA: { BG: '#7B1FA2', FONT: '#FFFFFF' }, // Dark Purple BG, White Font
        SAVINGS: { BG: '#1565C0', FONT: '#FFFFFF' }, // Blue BG, White Font
        DEFAULT: { BG: '#424242', FONT: '#FFFFFF' }, // Dark Gray BG, White Font
      },
      /** 
       * @property {object} UI General UI element colors.
       * Keys describe the UI element, values are hex color codes.
       */
      UI: {
        HEADER_BG: '#C62828', // Dark Red
        HEADER_FONT: '#FFFFFF', // White
        BORDER: '#FF8F00', // Orange
        INCOME_FONT: '#388E3C', // Green
        EXPENSE_FONT: '#D32F2F', // Red
        SAVINGS_FONT: '#1565C0', // Blue
        NEUTRAL_FONT: '#424242', // Dark Gray
        NET_BG: '#424242', // Dark Gray
        NET_FONT: '#FFFFFF', // White
      },
      /** @property {object} CHART Colors used in charts. */
      CHART: {
        /** @property {string[]} SERIES Array of hex color codes for chart series (e.g., Dark Red, Orange, Blue, Green, Purple, Dark Orange, Teal, Brown). */
        SERIES: [
          '#C62828', // Dark Red
          '#FF8F00', // Orange
          '#1565C0', // Blue
          '#2E7D32', // Dark Green
          '#6A1B9A', // Purple
          '#E64A19', // Dark Orange
          '#00695C', // Teal
          '#5D4037', // Brown
        ],
        /** @property {string} TITLE Color for chart titles (Dark Gray). */
        TITLE: '#424242',
        /** @property {string} TEXT Color for chart text (Dark Gray). */
        TEXT: '#424242',
      },
    },
    /** @property {object} CACHE Settings for application-level caching. */
    CACHE: {
      /** @property {boolean} ENABLED Flag to enable or disable caching. */
      ENABLED: true,
      /** 
       * @property {object} KEYS Defines keys used for storing cached data.
       * Keys are identifiers, values are the cache key strings.
       */
      KEYS: {
        CATEGORY_COMBINATIONS: 'finance_overview_categories',
        GROUPED_COMBINATIONS: 'finance_overview_grouped',
      },
      /** @property {number} EXPIRY_SECONDS Default cache expiry time in seconds (e.g., 21600 for 6 hours). */
      EXPIRY_SECONDS: 21600, // 6 hours
    },
    /** @property {object} LOCALE Localization settings. */
    LOCALE: {
      /** @property {string} CURRENCY_SYMBOL The currency symbol (e.g., '€', '$'). */
      CURRENCY_SYMBOL: '€', 
      /** @property {string} CURRENCY_LOCALE_CODE Locale code for currency formatting (e.g., '0' for default/system). */
      CURRENCY_LOCALE_CODE: '0', 
      /** @property {string} DATE_FORMAT Default date format string (e.g., 'yyyy-MM-dd'). */
      DATE_FORMAT: 'yyyy-MM-dd',
      /** 
       * @property {object} NUMBER_FORMATS Spreadsheet number formats for currency.
       * Keys describe the context, values are Google Sheets number format strings.
       */
      NUMBER_FORMATS: {
        CURRENCY_DEFAULT: '_-[$€-0]* #,##0_-;_-[RED][$€-0]* #,##0_-;* "-";_-@_-',
        CURRENCY_TOTAL_ROW: '_-[$€-0]* #,##0_-;_-[$€-0] (#,##0)_-;* "-";_-@_-' 
      }
    },
    /** @property {object} PERFORMANCE Performance-related settings. */
    PERFORMANCE: {
      /** @property {number} BATCH_SIZE Size of batches for operations like writing to sheets. */
      BATCH_SIZE: 50,
      /** @property {boolean} USE_BATCH_OPERATIONS Flag to enable or disable batch operations for better performance. */
      USE_BATCH_OPERATIONS: true,
    },
    /** @property {object} PLAID Plaid integration settings. */
    PLAID: {
      /** @property {string} API_URL The Plaid API base URL (sandbox by default). */
      API_URL: 'https://sandbox.plaid.com',
      /** 
       * @property {object} CATEGORY_MAP Mapping of Plaid categories to application transaction types.
       * Maps Plaid's primary category to the app's Type field.
       */
      CATEGORY_MAP: {
        'Food and Drink': 'Essentials',
        'Shops': 'Wants/Pleasure',
        'Travel': 'Wants/Pleasure',
        'Recreation': 'Wants/Pleasure',
        'Transfer': 'Savings',
        'Payment': 'Essentials',
        'Service': 'Essentials'
      }
    },
  };

  /**
   * Deeply merges properties from the source object into the target object.
   * @param {object} target The target object to merge into.
   * @param {object} source The source object to merge from.
   * @return {object} The modified target object.
   * @private
   */
  function mergeConfig(target, source) {
    const newTarget = JSON.parse(JSON.stringify(target)); // Create a deep copy of target
    for (const key in source) {
      if (Object.prototype.hasOwnProperty.call(source, key)) {
        if (source[key] instanceof Object && !(source[key] instanceof Array) && newTarget[key] instanceof Object && !(newTarget[key] instanceof Array)) {
          newTarget[key] = mergeConfig(newTarget[key], source[key]);
        } else {
          newTarget[key] = source[key];
        }
      }
    }
    return newTarget;
  }

  // User-specific configuration (can be overridden)
  let userConfig = {};

  // Public API
  return {
    /**
     * Retrieves the complete, merged configuration object (defaults merged with user overrides).
     * @returns {object} The fully merged configuration object.
     * @memberof FinancialPlanner.Config
     */
    get: function() {
      let currentConfig = JSON.parse(JSON.stringify(DEFAULT_CONFIG)); // Start with a fresh copy of defaults
      currentConfig = mergeConfig(currentConfig, userConfig);
      return currentConfig;
    },

    /**
     * Retrieves a specific section of the configuration.
     * @param {string} section - The key of the configuration section to retrieve (e.g., 'SHEETS', 'UI').
     * @returns {object} The configuration object for the specified section, or an empty object if not found.
     * @memberof FinancialPlanner.Config
     */
    getSection: function(section) {
      const config = this.get();
      return config[section] || {};
    },

    /**
     * Retrieves the sheet names configuration.
     * @returns {object} An object mapping internal sheet identifiers to their display names.
     * @memberof FinancialPlanner.Config
     */
    getSheetNames: function() {
      return this.getSection('SHEETS');
    },

    /**
     * Retrieves the transaction types configuration.
     * @returns {object} An object defining the different types of financial transactions.
     * @memberof FinancialPlanner.Config
     */
    getTransactionTypes: function() {
      return this.getSection('TRANSACTION_TYPES');
    },

    /**
     * Retrieves the target rates for financial planning (e.g., savings rate, expense ratios).
     * @returns {object} An object mapping financial categories to their target percentage rates.
     * @memberof FinancialPlanner.Config
     */
    getTargetRates: function() {
      return this.getSection('TARGET_RATES');
    },

    /**
     * Retrieves UI-specific configuration settings.
     * @returns {object} An object containing UI settings like column widths, toggle labels, etc.
     * @memberof FinancialPlanner.Config
     */
    getUI: function() {
      return this.getSection('UI');
    },

    /**
     * Retrieves color configurations for UI elements and charts.
     * @returns {object} An object defining color palettes for various application components.
     * @memberof FinancialPlanner.Config
     */
    getColors: function() {
      return this.getSection('COLORS');
    },

    /**
     * Retrieves locale-specific settings like currency symbol and date formats.
     * @returns {object} An object containing localization settings.
     * @memberof FinancialPlanner.Config
     */
    getLocale: function() {
      return this.getSection('LOCALE');
    },

    /**
     * Updates the user-specific configuration by merging new settings.
     * @param {object} configToUpdate - An object containing configuration settings to merge into the user config.
     * @memberof FinancialPlanner.Config
     */
    update: function(configToUpdate) {
      userConfig = mergeConfig(userConfig, configToUpdate);
    },

    /**
     * Resets the user-specific configuration to an empty object, effectively reverting to default settings.
     * @memberof FinancialPlanner.Config
     */
    reset: function() {
      userConfig = {};
    }
  };
})();
