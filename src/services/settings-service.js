/**
 * @fileoverview Settings Service Module for Financial Planning Tools.
 * Manages user preferences and application settings by storing them in a dedicated,
 * hidden sheet within the spreadsheet. Provides methods to get, set, and manage various types of settings.
 * @module services/settings-service
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Settings Service - Manages user preferences stored in a hidden settings sheet.
 * @namespace FinancialPlanner.SettingsService
 */
FinancialPlanner.SettingsService = (function() {
  /**
   * Retrieves or creates the settings sheet.
   * The sheet is hidden by default if created.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The settings sheet object.
   * @private
   */
  function getSettingsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = FinancialPlanner.Config.getSheetNames().SETTINGS;
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange('A1:B1').setValues([['Preference', 'Value']]).setFontWeight('bold');
      sheet.hideSheet();
    }
    return sheet;
  }

  /**
   * Finds a preference row and its value in the settings sheet.
   * @param {string} key - The preference key to find.
   * @returns {{row: number, value: *}|null} An object containing the 1-based row number
   *   and the value of the preference, or null if the key is not found.
   * @private
   */
  function findPreference(key) {
    const sheet = getSettingsSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return { row: i + 1, value: data[i][1] };
      }
    }
    return null;
  }

  // Public API
  return {
    /**
     * Retrieves the value of a preference by its key.
     * @param {string} key - The key of the preference to retrieve.
     * @param {*} [defaultValue] - The value to return if the key is not found.
     * @returns {*} The stored value of the preference, or the `defaultValue` if not found or an error occurs.
     * @memberof FinancialPlanner.SettingsService
     */
    getValue: function(key, defaultValue) {
      try {
        const preference = findPreference(key);
        return preference ? preference.value : defaultValue;
      } catch (error) {
        FinancialPlanner.ErrorService.handle(
          FinancialPlanner.ErrorService.create('Error getting setting value for key: ' + key, { originalError: error.toString(), severity: 'medium' }),
          'Failed to get setting: ' + key
        );
        return defaultValue;
      }
    },

    /**
     * Sets the value of a preference. If the key exists, it updates the value.
     * If the key does not exist, it adds a new row for the preference.
     * @param {string} key - The key of the preference to set.
     * @param {*} value - The value to store for the preference.
     * @memberof FinancialPlanner.SettingsService
     */
    setValue: function(key, value) {
      try {
        const sheet = getSettingsSheet();
        const preference = findPreference(key);
        if (preference) {
          sheet.getRange(preference.row, 2).setValue(value);
        } else {
          const lastRow = Math.max(1, sheet.getLastRow());
          sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
        }
      } catch (error) {
        FinancialPlanner.ErrorService.handle(
          FinancialPlanner.ErrorService.create('Error setting setting value for key: ' + key, { originalError: error.toString(), valueToSet: value, severity: 'high' }),
          'Failed to set setting: ' + key
        );
      }
    },

    /**
     * Retrieves a preference value and coerces it to a boolean.
     * Handles 'true'/'false' strings and 0/1 numbers.
     * @param {string} key - The key of the preference.
     * @param {boolean} [defaultValue=false] - The default boolean value if the key is not found or cannot be coerced.
     * @returns {boolean} The boolean value of the preference.
     * @memberof FinancialPlanner.SettingsService
     */
    getBooleanValue: function(key, defaultValue) {
      defaultValue = defaultValue !== undefined ? defaultValue : false;
      const value = this.getValue(key, defaultValue);
      if (typeof value === 'boolean') return value;
      if (value === 'true' || value === 1 || value === '1') return true;
      if (value === 'false' || value === 0 || value === '0') return false;
      return !!defaultValue;
    },

    /**
     * Retrieves a preference value and coerces it to a number.
     * @param {string} key - The key of the preference.
     * @param {number} [defaultValue=0] - The default numeric value if the key is not found or cannot be parsed.
     * @returns {number} The numeric value of the preference.
     * @memberof FinancialPlanner.SettingsService
     */
    getNumericValue: function(key, defaultValue) {
      defaultValue = defaultValue !== undefined ? defaultValue : 0;
      const value = this.getValue(key, defaultValue);
      if (typeof value === 'number') return value;
      const parsed = parseFloat(value);
      return isNaN(parsed) ? (typeof defaultValue === 'number' ? defaultValue : 0) : parsed;
    },

    /**
     * Toggles a boolean preference value.
     * Retrieves the current boolean value, inverts it, and saves the new value.
     * @param {string} key - The key of the boolean preference to toggle.
     * @param {boolean} [defaultValue=false] - The default value to assume if the preference is not yet set.
     * @returns {boolean} The new boolean value after toggling.
     * @memberof FinancialPlanner.SettingsService
     */
    toggleBooleanValue: function(key, defaultValue) {
      defaultValue = defaultValue !== undefined ? defaultValue : false;
      const currentValue = this.getBooleanValue(key, defaultValue);
      const newValue = !currentValue;
      this.setValue(key, newValue);
      return newValue;
    },

    /**
     * Toggles the 'ShowSubCategories' preference.
     * Defaults to true if not set.
     * @returns {boolean} The new value of the 'ShowSubCategories' preference.
     * @memberof FinancialPlanner.SettingsService
     */
    toggleShowSubCategories: function() {
      return this.toggleBooleanValue('ShowSubCategories', true);
    },

    /**
     * Gets the current value of the 'ShowSubCategories' preference.
     * Defaults to true if not set.
     * @returns {boolean} The current boolean value of the 'ShowSubCategories' preference.
     * @memberof FinancialPlanner.SettingsService
     */
    getShowSubCategories: function() {
      return this.getBooleanValue('ShowSubCategories', true);
    },

    /**
     * Sets the value of the 'ShowSubCategories' preference.
     * @param {boolean} value - The boolean value to set for 'ShowSubCategories'.
     * @memberof FinancialPlanner.SettingsService
     */
    setShowSubCategories: function(value) {
      this.setValue('ShowSubCategories', typeof value === 'boolean' ? value : true);
    },

    /**
     * Gets the current Plaid environment setting.
     * @returns {string} The current environment ('sandbox' or 'production'). Defaults to 'sandbox'.
     * @memberof FinancialPlanner.SettingsService
     */
    getPlaidEnvironment: function() {
      return this.getValue('PlaidEnvironment', 'sandbox');
    },

    /**
     * Sets the Plaid environment preference.
     * @param {string} environment - The environment to set ('sandbox' or 'production').
     * @throws {Error} If the environment is invalid.
     * @memberof FinancialPlanner.SettingsService
     */
    setPlaidEnvironment: function(environment) {
      if (environment !== 'sandbox' && environment !== 'production') {
        throw FinancialPlanner.ErrorService.create(
          'Invalid environment. Must be "sandbox" or "production"',
          { severity: 'high', providedValue: environment }
        );
      }
      this.setValue('PlaidEnvironment', environment);
    },

    /**
     * Gets the stored SaltEdge customer ID.
     * @returns {string|null} SaltEdge customer ID or null if not set.
     * @memberof FinancialPlanner.SettingsService
     */
    getSaltEdgeCustomerId: function() {
      return this.getValue('SaltEdgeCustomerId', null);
    },

    /**
     * Sets the SaltEdge customer ID preference.
     * @param {string} customerId - The SaltEdge customer ID to store.
     * @memberof FinancialPlanner.SettingsService
     */
    setSaltEdgeCustomerId: function(customerId) {
      this.setValue('SaltEdgeCustomerId', customerId);
    },

    /**
     * Retrieves all preferences stored in the settings sheet as an object.
     * @returns {Object<string, *>} An object where keys are preference names and values are their stored values.
     *   Returns an empty object if an error occurs.
     * @memberof FinancialPlanner.SettingsService
     */
    getAllPreferences: function() {
      try {
        const sheet = getSettingsSheet();
        const data = sheet.getDataRange().getValues();
        const preferences = {};
        for (let i = 1; i < data.length; i++) {
          if (data[i] && data[i][0] != null && data[i][0] !== '') {
            preferences[data[i][0]] = data[i][1];
          }
        }
        return preferences;
      } catch (error) {
        FinancialPlanner.ErrorService.handle(
          FinancialPlanner.ErrorService.create('Error getting all preferences', { originalError: error.toString(), severity: 'medium' }),
          'Failed to retrieve all settings.'
        );
        return {};
      }
    },

    /**
     * Clears all preferences from the settings sheet, effectively resetting them.
     * The header row is preserved. Shows a success notification via UIService.
     * @memberof FinancialPlanner.SettingsService
     */
    resetAllPreferences: function() {
      try {
        const sheet = getSettingsSheet();
        if (sheet.getLastRow() > 1) {
          sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
        }
        FinancialPlanner.UIService.showSuccessNotification('All preferences have been reset.');
      } catch (error) {
        FinancialPlanner.ErrorService.handle(
          FinancialPlanner.ErrorService.create('Error resetting all preferences', { originalError: error.toString(), severity: 'high' }),
          'Failed to reset settings.'
        );
      }
    }
  };
})();
