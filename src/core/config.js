/**
 * @fileoverview Configuration Module for Financial Planning Tools.
 * Provides centralized configuration settings for the application.
 * This module is designed to be instantiated by the 00_module_loader.js.
 */

// This self-executing function encapsulates the module's private scope
// and returns the constructor function. This pattern is used to keep
// DEFAULT_CONFIG and mergeConfig private to the ConfigModule instances.
// eslint-disable-next-line no-unused-vars
const ConfigModule = (function() {
  // Default configuration - kept in the closure to be "private"
  const DEFAULT_CONFIG = {
    SHEETS: {
      OVERVIEW: 'Overview',
      TRANSACTIONS: 'Transactions',
      DROPDOWNS: 'Dropdowns',
      ERROR_LOG: 'Error Log',
      ANALYSIS: 'Analysis',
      SETTINGS: 'Settings',
    },
    TRANSACTION_TYPES: {
      INCOME: 'Income',
      ESSENTIALS: 'Essentials',
      WANTS: 'Wants/Pleasure',
      EXTRA: 'Extra',
      SAVINGS: 'Savings',
    },
    TYPE_ORDER: ['Income', 'Essentials', 'Wants/Pleasure', 'Extra', 'Savings'],
    EXPENSE_TYPES: ['Essentials', 'Wants/Pleasure', 'Extra'],
    TARGET_RATES: {
      ESSENTIALS: 0.5,
      WANTS: 0.2,
      EXTRA: 0.1,
      SAVINGS: 0.2,
      DEFAULT: 0.2,
    },
    HEADERS: [
      'Type', 'Category', 'Sub-Category', 'Shared?',
      'Jan-24', 'Feb-24', 'Mar-24', 'Apr-24',
      'May-24', 'Jun-24', 'Jul-24', 'Aug-24',
      'Sep-24', 'Oct-24', 'Nov-24', 'Dec-24', 'Total', 'Average',
    ],
    UI: {
      SUBCATEGORY_TOGGLE: {
        LABEL_CELL: 'S1',
        CHECKBOX_CELL: 'T1',
        LABEL_TEXT: 'Show Sub-Categories',
        NOTE_TEXT: 'Toggle to show or hide sub-categories in the overview sheet',
      },
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
    COLORS: {
      TYPE_HEADERS: {
        INCOME: { BG: '#2E7D32', FONT: '#FFFFFF' },
        ESSENTIALS: { BG: '#1976D2', FONT: '#FFFFFF' },
        WANTS_PLEASURE: { BG: '#FFA000', FONT: '#FFFFFF' },
        EXTRA: { BG: '#7B1FA2', FONT: '#FFFFFF' },
        SAVINGS: { BG: '#1565C0', FONT: '#FFFFFF' },
        DEFAULT: { BG: '#424242', FONT: '#FFFFFF' },
      },
      UI: {
        HEADER_BG: '#C62828',
        HEADER_FONT: '#FFFFFF',
        BORDER: '#FF8F00',
        INCOME_FONT: '#388E3C',
        EXPENSE_FONT: '#D32F2F',
        SAVINGS_FONT: '#1565C0',
        NEUTRAL_FONT: '#424242',
        NET_BG: '#424242',
        NET_FONT: '#FFFFFF',
      },
      CHART: {
        SERIES: [
          '#C62828', '#FF8F00', '#1565C0', '#2E7D32',
          '#6A1B9A', '#E64A19', '#00695C', '#5D4037',
        ],
        TITLE: '#424242',
        TEXT: '#424242',
      },
    },
    CACHE: {
      ENABLED: true,
      KEYS: {
        CATEGORY_COMBINATIONS: 'finance_overview_categories',
        GROUPED_COMBINATIONS: 'finance_overview_grouped',
      },
      EXPIRY_SECONDS: 21600, // 6 hours
    },
    LOCALE: {
      CURRENCY_SYMBOL: '€',
      CURRENCY_LOCALE: '0',
      DATE_FORMAT: 'yyyy-MM-dd',
    },
    PERFORMANCE: {
      BATCH_SIZE: 50,
      USE_BATCH_OPERATIONS: true,
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

  /**
   * Constructor for the ConfigModule.
   * Initializes user-specific configuration.
   * @constructor
   */
  function ConfigModuleConstructor() {
    this.userConfig = {};
    // All public methods will be attached to 'this' or its prototype.
  }

  ConfigModuleConstructor.prototype.get = function() {
    let currentConfig = JSON.parse(JSON.stringify(DEFAULT_CONFIG)); // Start with a fresh copy of defaults
    currentConfig = mergeConfig(currentConfig, this.userConfig);
    return currentConfig;
  };

  ConfigModuleConstructor.prototype.getSection = function(section) {
    const config = this.get();
    return config[section] || {};
  };

  ConfigModuleConstructor.prototype.getSheetNames = function() {
    return this.getSection('SHEETS');
  };

  ConfigModuleConstructor.prototype.getTransactionTypes = function() {
    return this.getSection('TRANSACTION_TYPES');
  };

  ConfigModuleConstructor.prototype.getTargetRates = function() {
    return this.getSection('TARGET_RATES');
  };

  ConfigModuleConstructor.prototype.getUI = function() {
    return this.getSection('UI');
  };

  ConfigModuleConstructor.prototype.getColors = function() {
    return this.getSection('COLORS');
  };

  ConfigModuleConstructor.prototype.getLocale = function() {
    return this.getSection('LOCALE');
  };

  ConfigModuleConstructor.prototype.update = function(configToUpdate) {
    this.userConfig = mergeConfig(this.userConfig, configToUpdate);
  };

  ConfigModuleConstructor.prototype.reset = function() {
    this.userConfig = {};
  };

  return ConfigModuleConstructor;
})();
