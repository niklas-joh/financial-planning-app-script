/**
 * Financial Planning Tools - Settings Service
 * 
 * This file provides a centralized service for managing user preferences and settings.
 * It helps ensure consistent settings management across the application.
 */

// Create the SettingsService module within the FinancialPlanner namespace
FinancialPlanner.SettingsService = (function(config, utils) {
  // Private variables and functions
  
  /**
   * Gets the settings sheet, creating it if it doesn't exist
   * @return {SpreadsheetApp.Sheet} The settings sheet
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
   * Finds a preference in the settings sheet
   * @param {String} key - The preference key to find
   * @return {Object} Object containing the row index and current value, or null if not found
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
     * Gets a preference value
     * @param {String} key - The preference key
     * @param {any} defaultValue - The default value to return if the preference doesn't exist
     * @return {any} The preference value, or the default value if not found
     */
    getValue: function(key, defaultValue) {
      const preference = findPreference(key);
      
      if (preference) {
        return preference.value;
      }
      
      return defaultValue;
    },
    
    /**
     * Sets a preference value
     * @param {String} key - The preference key
     * @param {any} value - The preference value
     */
    setValue: function(key, value) {
      const sheet = getSettingsSheet();
      const preference = findPreference(key);
      
      if (preference) {
        // Update existing preference
        sheet.getRange(preference.row, 2).setValue(value);
      } else {
        // Add new preference
        const lastRow = Math.max(1, sheet.getLastRow());
        sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
      }
    },
    
    /**
     * Gets a boolean preference value
     * @param {String} key - The preference key
     * @param {Boolean} defaultValue - The default value to return if the preference doesn't exist
     * @return {Boolean} The preference value as a boolean
     */
    getBooleanValue: function(key, defaultValue) {
      const value = this.getValue(key, defaultValue);
      
      // Convert to boolean if it's not already
      if (typeof value === 'boolean') {
        return value;
      }
      
      // Convert string values
      if (value === 'true') return true;
      if (value === 'false') return false;
      
      // Convert numeric values
      if (value === 1 || value === '1') return true;
      if (value === 0 || value === '0') return false;
      
      // Default to the provided default value
      return defaultValue;
    },
    
    /**
     * Gets a numeric preference value
     * @param {String} key - The preference key
     * @param {Number} defaultValue - The default value to return if the preference doesn't exist
     * @return {Number} The preference value as a number
     */
    getNumericValue: function(key, defaultValue) {
      const value = this.getValue(key, defaultValue);
      
      // Convert to number if it's not already
      if (typeof value === 'number') {
        return value;
      }
      
      // Try to parse as a number
      const parsed = parseFloat(value);
      
      // Return the parsed value if it's a valid number, otherwise return the default
      return isNaN(parsed) ? defaultValue : parsed;
    },
    
    /**
     * Toggles a boolean preference value
     * @param {String} key - The preference key
     * @param {Boolean} defaultValue - The default value to use if the preference doesn't exist
     * @return {Boolean} The new value after toggling
     */
    toggleBooleanValue: function(key, defaultValue) {
      const currentValue = this.getBooleanValue(key, defaultValue);
      const newValue = !currentValue;
      
      this.setValue(key, newValue);
      
      return newValue;
    },
    
    /**
     * Toggles the display of sub-categories in the overview
     * @return {Boolean} The new value after toggling
     */
    toggleShowSubCategories: function() {
      return this.toggleBooleanValue("ShowSubCategories", true);
    },
    
    /**
     * Gets whether to show sub-categories in the overview
     * @return {Boolean} True if sub-categories should be shown, false otherwise
     */
    getShowSubCategories: function() {
      return this.getBooleanValue("ShowSubCategories", true);
    },
    
    /**
     * Sets whether to show sub-categories in the overview
     * @param {Boolean} value - Whether to show sub-categories
     */
    setShowSubCategories: function(value) {
      this.setValue("ShowSubCategories", value);
    },
    
    /**
     * Gets all preferences as an object
     * @return {Object} Object containing all preferences
     */
    getAllPreferences: function() {
      const sheet = getSettingsSheet();
      const data = sheet.getDataRange().getValues();
      const preferences = {};
      
      // Start from row 1 (index 0) to skip the header
      for (let i = 1; i < data.length; i++) {
        preferences[data[i][0]] = data[i][1];
      }
      
      return preferences;
    },
    
    /**
     * Resets all preferences to their default values
     */
    resetAllPreferences: function() {
      const sheet = getSettingsSheet();
      
      // Clear all rows except the header
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
      }
    }
  };
})(FinancialPlanner.Config, FinancialPlanner.Utils);
