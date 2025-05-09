/**
 * @fileoverview Settings Service Module for Financial Planning Tools.
 * Manages user preferences and application settings.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const SettingsServiceModule = (function() {
  /**
   * Constructor for the SettingsServiceModule.
   * @param {object} configInstance - An instance of ConfigModule.
   * @param {object} uiServiceInstance - An instance of UIServiceModule.
   * @param {object} errorServiceInstance - An instance of ErrorServiceModule.
   * @constructor
   */
  function SettingsServiceModuleConstructor(configInstance, uiServiceInstance, errorServiceInstance) {
    this.config = configInstance;
    this.uiService = uiServiceInstance;
    this.errorService = errorServiceInstance;
    // FinancialPlanner.Utils is assumed to be globally available or refactored separately.
  }

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

  SettingsServiceModuleConstructor.prototype.getValue = function(key, defaultValue) {
    try {
      const preference = this._findPreference(key);
      return preference ? preference.value : defaultValue;
    } catch (error) {
      this.errorService.handle(this.errorService.create(`Error getting setting value for key: ${key}`, { originalError: error.toString(), severity: 'medium' }), `Failed to get setting: ${key}`);
      return defaultValue;
    }
  };

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

  SettingsServiceModuleConstructor.prototype.getBooleanValue = function(key, defaultValue = false) {
    const value = this.getValue(key, defaultValue);
    if (typeof value === 'boolean') return value;
    if (value === 'true' || value === 1 || value === '1') return true;
    if (value === 'false' || value === 0 || value === '0') return false;
    return !!defaultValue;
  };

  SettingsServiceModuleConstructor.prototype.getNumericValue = function(key, defaultValue = 0) {
    const value = this.getValue(key, defaultValue);
    if (typeof value === 'number') return value;
    const parsed = parseFloat(value);
    return isNaN(parsed) ? (typeof defaultValue === 'number' ? defaultValue : 0) : parsed;
  };

  SettingsServiceModuleConstructor.prototype.toggleBooleanValue = function(key, defaultValue = false) {
    const currentValue = this.getBooleanValue(key, defaultValue);
    const newValue = !currentValue;
    this.setValue(key, newValue);
    return newValue;
  };

  SettingsServiceModuleConstructor.prototype.toggleShowSubCategories = function() {
    return this.toggleBooleanValue('ShowSubCategories', true);
  };

  SettingsServiceModuleConstructor.prototype.getShowSubCategories = function() {
    return this.getBooleanValue('ShowSubCategories', true);
  };

  SettingsServiceModuleConstructor.prototype.setShowSubCategories = function(value) {
    this.setValue('ShowSubCategories', typeof value === 'boolean' ? value : true);
  };

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
