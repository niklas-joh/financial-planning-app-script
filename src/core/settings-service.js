/**
 * Financial Planning Tools - Settings Service
 * 
 * This file provides a centralized service for managing user preferences and settings.
 * It helps ensure consistent settings management across the application.
 */

/**
 * @namespace FinancialPlanner.SettingsService
 * @param {FinancialPlanner.Config} config - The configuration service.
 * @param {FinancialPlanner.Utils} utils - The utility service.
 * @param {FinancialPlanner.UIService} uiService - The UI service.
 * @param {FinancialPlanner.ErrorService} errorService - The error handling service.
 */
FinancialPlanner.SettingsService = (function(config, utils, uiService, errorService) {
  // Private variables and functions
  
  /**
   * Retrieves the dedicated 'Settings' sheet from the active spreadsheet.
   * If the sheet does not exist, it creates and initializes it with a header row ("Preference", "Value")
   * and then hides the sheet.
   * @return {GoogleAppsScript.Spreadsheet.Sheet} The settings sheet object.
   * @private
   */
  function getSettingsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = config.getSheetNames().SETTINGS;
    
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      // Create a new settings sheet
      sheet = ss.insertSheet(sheetName);
      
      // Set up the header row
      sheet.getRange("A1:B1").setValues([["Preference", "Value"]]);
      sheet.getRange("A1:B1").setFontWeight("bold");
      
      // Hide the sheet (it's just for storing settings)
      sheet.hideSheet();
    }
    
    return sheet;
  }
  
  /**
   * Searches for a preference key in the settings sheet.
   * @param {string} key - The preference key to find (e.g., "ShowSubCategories").
   * @return {{row: number, value: any} | null} An object containing the 1-based row index and the current value
   *                                            if the key is found; otherwise, null.
   * @private
   */
  function findPreference(key) {
    const sheet = getSettingsSheet();
    const data = sheet.getDataRange().getValues();
    
    // Start from row 1 (index 0) to skip the header
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return {
          row: i + 1, // Convert to 1-based index
          value: data[i][1]
        };
      }
    }
    
    return null;
  }
  
  // Public API
  return {
    /**
     * Retrieves the value of a specific preference.
     * If the preference key is not found, it returns the `defaultValue`.
     * Errors are handled by `errorService`.
     * @param {string} key - The unique key for the preference (e.g., "UserTheme").
     * @param {any} [defaultValue] - The value to return if the key is not found.
     * @return {any} The stored preference value, or `defaultValue`.
     *
     * @example
     * const theme = FinancialPlanner.SettingsService.getValue("UserTheme", "light");
     * const itemsPerPage = FinancialPlanner.SettingsService.getValue("ItemsPerPage", 10);
     */
    getValue: function(key, defaultValue) {
      try {
        const preference = findPreference(key);
        return preference ? preference.value : defaultValue;
      } catch (error) {
        errorService.handle(errorService.create(`Error getting setting value for key: ${key}`, { originalError: error.toString() }), `Failed to get setting: ${key}`);
        return defaultValue;
      }
    },
    
    /**
     * Sets or updates the value of a specific preference.
     * If the preference key exists, its value is updated. If not, a new preference entry is created.
     * Errors are handled by `errorService`.
     * @param {string} key - The unique key for the preference.
     * @param {any} value - The value to store for the preference.
     * @return {void}
     *
     * @example
     * FinancialPlanner.SettingsService.setValue("UserTheme", "dark");
     * FinancialPlanner.SettingsService.setValue("NotificationsEnabled", true);
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
        errorService.handle(errorService.create(`Error setting setting value for key: ${key}`, { originalError: error.toString(), valueToSet: value }), `Failed to set setting: ${key}`);
      }
    },
    
    /**
     * Retrieves a preference value and coerces it to a boolean.
     * Handles string "true"/"false" and numeric 1/0 as boolean.
     * @param {string} key - The preference key.
     * @param {boolean} [defaultValue=false] - The default boolean value if the key is not found or cannot be coerced.
     * @return {boolean} The preference value as a boolean.
     *
     * @example
     * const showTips = FinancialPlanner.SettingsService.getBooleanValue("ShowStartupTips", true);
     */
    getBooleanValue: function(key, defaultValue) {
      const value = this.getValue(key, defaultValue);
      if (typeof value === 'boolean') return value;
      if (value === 'true' || value === 1 || value === '1') return true;
      if (value === 'false' || value === 0 || value === '0') return false;
      return !!defaultValue; // Ensure defaultValue is also coerced if not boolean
    },
    
    /**
     * Retrieves a preference value and coerces it to a number.
     * @param {string} key - The preference key.
     * @param {number} [defaultValue=0] - The default numeric value if the key is not found or cannot be parsed.
     * @return {number} The preference value as a number.
     *
     * @example
     * const maxItems = FinancialPlanner.SettingsService.getNumericValue("MaxDashboardItems", 5);
     */
    getNumericValue: function(key, defaultValue) {
      const value = this.getValue(key, defaultValue);
      if (typeof value === 'number') return value;
      const parsed = parseFloat(value);
      return isNaN(parsed) ? (typeof defaultValue === 'number' ? defaultValue : 0) : parsed;
    },
    
    /**
     * Toggles a boolean preference value. If the key doesn't exist, it uses `defaultValue` to determine the initial state before toggling.
     * The new (toggled) value is then saved.
     * @param {string} key - The preference key.
     * @param {boolean} [defaultValue=false] - The default value to assume if the preference doesn't exist.
     * @return {boolean} The new value after toggling.
     *
     * @example
     * const newNotificationState = FinancialPlanner.SettingsService.toggleBooleanValue("EnableNotifications", true);
     * // If it was true, it's now false. If it was false or unset (defaulting to true), it's now false.
     */
    toggleBooleanValue: function(key, defaultValue) {
      const currentValue = this.getBooleanValue(key, defaultValue);
      const newValue = !currentValue;
      this.setValue(key, newValue);
      return newValue;
    },
    
    /**
     * Toggles the "ShowSubCategories" preference. Defaults to true if not set.
     * @return {boolean} The new value of the "ShowSubCategories" preference.
     */
    toggleShowSubCategories: function() {
      return this.toggleBooleanValue("ShowSubCategories", true);
    },
    
    /**
     * Gets the current "ShowSubCategories" preference. Defaults to true if not set.
     * @return {boolean} True if sub-categories should be shown, false otherwise.
     */
    getShowSubCategories: function() {
      return this.getBooleanValue("ShowSubCategories", true);
    },
    
    /**
     * Sets the "ShowSubCategories" preference.
     * @param {boolean} value - Whether to show sub-categories.
     * @return {void}
     */
    setShowSubCategories: function(value) {
      this.setValue("ShowSubCategories", typeof value === 'boolean' ? value : true);
    },
    
    /**
     * Retrieves all preferences stored in the settings sheet as an object.
     * @return {object} An object where keys are preference names and values are their stored values.
     *                  Returns an empty object if an error occurs.
     *
     * @example
     * const allSettings = FinancialPlanner.SettingsService.getAllPreferences();
     * console.log(allSettings["UserTheme"]);
     */
    getAllPreferences: function() {
      try {
        const sheet = getSettingsSheet();
        const data = sheet.getDataRange().getValues();
        const preferences = {};
        for (let i = 1; i < data.length; i++) { // Start from row 1 to skip header
          if (data[i] && data[i][0] != null && data[i][0] !== "") {
            preferences[data[i][0]] = data[i][1];
          }
        }
        return preferences;
      } catch (error) {
        errorService.handle(errorService.create('Error getting all preferences', { originalError: error.toString() }), "Failed to retrieve all settings.");
        return {};
      }
    },
    
    /**
     * Resets all preferences by clearing all entries in the settings sheet except the header row.
     * Shows a success notification upon completion. Errors are handled by `errorService`.
     * @return {void}
     *
     * @example
     * FinancialPlanner.SettingsService.resetAllPreferences();
     */
    resetAllPreferences: function() {
      try {
        const sheet = getSettingsSheet();
        if (sheet.getLastRow() > 1) {
          sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
        }
        uiService.showSuccessNotification("All preferences have been reset.");
      } catch (error) {
        errorService.handle(errorService.create('Error resetting all preferences', { originalError: error.toString() }), "Failed to reset settings.");
      }
    }
  };
})(FinancialPlanner.Config, FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.ErrorService);
