/**
 * @fileoverview Settings Service Module for Financial Planning Tools.
 * Manages user preferences and application settings by storing them in a dedicated,
 * hidden sheet within the spreadsheet. Provides methods to get, set, and manage various types of settings.
 * This module is designed to be instantiated by `00_module_loader.js`.
 * @module services/settings-service
 */

/**
 * IIFE to encapsulate the SettingsServiceModule logic.
 * @returns {function} The SettingsServiceModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const SettingsServiceModule = (function() {
  /**
   * Constructor for the SettingsServiceModule.
   * @param {ConfigModule} configInstance - An instance of ConfigModule.
   * @param {UIServiceModule} uiServiceInstance - An instance of UIServiceModule.
   * @param {ErrorServiceModule} errorServiceInstance - An instance of ErrorServiceModule.
   * @constructor
   * @alias SettingsServiceModule
   * @memberof module:services/settings-service
   */
  function SettingsServiceModuleConstructor(configInstance, uiServiceInstance, errorServiceInstance) {
    /**
     * Instance of ConfigModule.
     * @type {ConfigModule}
     * @private
     */
    this.config = configInstance;
    /**
     * Instance of UIServiceModule.
     * @type {UIServiceModule}
     * @private
     */
    this.uiService = uiServiceInstance;
    /**
     * Instance of ErrorServiceModule.
     * @type {ErrorServiceModule}
     * @private
     */
    this.errorService = errorServiceInstance;
    // FinancialPlanner.Utils is assumed to be globally available or refactored separately.
  }

  /**
   * Retrieves or creates the settings sheet.
   * The sheet is hidden by default if created.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The settings sheet object.
   * @private
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype._getSettingsSheet = function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = this.config.getSheetNames().SETTINGS;
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange('A1:B1').setValues([['Preference', 'Value']]).setFontWeight('bold');
      sheet.hideSheet();
    }
    return sheet;
  };

  /**
   * Finds a preference row and its value in the settings sheet.
   * @param {string} key - The preference key to find.
   * @returns {{row: number, value: *}|null} An object containing the 1-based row number
   *   and the value of the preference, or null if the key is not found.
   * @private
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype._findPreference = function(key) {
    const sheet = this._getSettingsSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return { row: i + 1, value: data[i][1] };
      }
    }
    return null;
  };

  /**
   * Retrieves the value of a preference by its key.
   * @param {string} key - The key of the preference to retrieve.
   * @param {*} [defaultValue] - The value to return if the key is not found.
   * @returns {*} The stored value of the preference, or the `defaultValue` if not found or an error occurs.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.getValue = function(key, defaultValue) {
    try {
      const preference = this._findPreference(key);
      return preference ? preference.value : defaultValue;
    } catch (error) {
      this.errorService.handle(this.errorService.create(`Error getting setting value for key: ${key}`, { originalError: error.toString(), severity: 'medium' }), `Failed to get setting: ${key}`);
      return defaultValue;
    }
  };

  /**
   * Sets the value of a preference. If the key exists, it updates the value.
   * If the key does not exist, it adds a new row for the preference.
   * @param {string} key - The key of the preference to set.
   * @param {*} value - The value to store for the preference.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.setValue = function(key, value) {
    try {
      const sheet = this._getSettingsSheet();
      const preference = this._findPreference(key);
      if (preference) {
        sheet.getRange(preference.row, 2).setValue(value);
      } else {
        const lastRow = Math.max(1, sheet.getLastRow());
        sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
      }
    } catch (error) {
      this.errorService.handle(this.errorService.create(`Error setting setting value for key: ${key}`, { originalError: error.toString(), valueToSet: value, severity: 'high' }), `Failed to set setting: ${key}`);
    }
  };

  /**
   * Retrieves a preference value and coerces it to a boolean.
   * Handles 'true'/'false' strings and 0/1 numbers.
   * @param {string} key - The key of the preference.
   * @param {boolean} [defaultValue=false] - The default boolean value if the key is not found or cannot be coerced.
   * @returns {boolean} The boolean value of the preference.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.getBooleanValue = function(key, defaultValue = false) {
    const value = this.getValue(key, defaultValue);
    if (typeof value === 'boolean') return value;
    if (value === 'true' || value === 1 || value === '1') return true;
    if (value === 'false' || value === 0 || value === '0') return false;
    return !!defaultValue;
  };

  /**
   * Retrieves a preference value and coerces it to a number.
   * @param {string} key - The key of the preference.
   * @param {number} [defaultValue=0] - The default numeric value if the key is not found or cannot be parsed.
   * @returns {number} The numeric value of the preference.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.getNumericValue = function(key, defaultValue = 0) {
    const value = this.getValue(key, defaultValue);
    if (typeof value === 'number') return value;
    const parsed = parseFloat(value);
    return isNaN(parsed) ? (typeof defaultValue === 'number' ? defaultValue : 0) : parsed;
  };

  /**
   * Toggles a boolean preference value.
   * Retrieves the current boolean value, inverts it, and saves the new value.
   * @param {string} key - The key of the boolean preference to toggle.
   * @param {boolean} [defaultValue=false] - The default value to assume if the preference is not yet set.
   * @returns {boolean} The new boolean value after toggling.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.toggleBooleanValue = function(key, defaultValue = false) {
    const currentValue = this.getBooleanValue(key, defaultValue);
    const newValue = !currentValue;
    this.setValue(key, newValue);
    return newValue;
  };

  /**
   * Toggles the 'ShowSubCategories' preference.
   * Defaults to true if not set.
   * @returns {boolean} The new value of the 'ShowSubCategories' preference.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.toggleShowSubCategories = function() {
    return this.toggleBooleanValue('ShowSubCategories', true);
  };

  /**
   * Gets the current value of the 'ShowSubCategories' preference.
   * Defaults to true if not set.
   * @returns {boolean} The current boolean value of the 'ShowSubCategories' preference.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.getShowSubCategories = function() {
    return this.getBooleanValue('ShowSubCategories', true);
  };

  /**
   * Sets the value of the 'ShowSubCategories' preference.
   * @param {boolean} value - The boolean value to set for 'ShowSubCategories'.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.setShowSubCategories = function(value) {
    this.setValue('ShowSubCategories', typeof value === 'boolean' ? value : true);
  };

  /**
   * Retrieves all preferences stored in the settings sheet as an object.
   * @returns {Object<string, *>} An object where keys are preference names and values are their stored values.
   *   Returns an empty object if an error occurs.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.getAllPreferences = function() {
    try {
      const sheet = this._getSettingsSheet();
      const data = sheet.getDataRange().getValues();
      const preferences = {};
      for (let i = 1; i < data.length; i++) {
        if (data[i] && data[i][0] != null && data[i][0] !== '') {
          preferences[data[i][0]] = data[i][1];
        }
      }
      return preferences;
    } catch (error) {
      this.errorService.handle(this.errorService.create('Error getting all preferences', { originalError: error.toString(), severity: 'medium' }), 'Failed to retrieve all settings.');
      return {};
    }
  };

  /**
   * Clears all preferences from the settings sheet, effectively resetting them.
   * The header row is preserved. Shows a success notification via UIService.
   * @memberof SettingsServiceModule
   */
  SettingsServiceModuleConstructor.prototype.resetAllPreferences = function() {
    try {
      const sheet = this._getSettingsSheet();
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
      }
      this.uiService.showSuccessNotification('All preferences have been reset.');
    } catch (error) {
      this.errorService.handle(this.errorService.create('Error resetting all preferences', { originalError: error.toString(), severity: 'high' }), 'Failed to reset settings.');
    }
  };

  return SettingsServiceModuleConstructor;
})();
