/**
 * Financial Planning Tools - Settings and Configuration
 * 
 * This file contains functions for managing user settings and preferences
 */

/**
 * Gets a user preference value from the Settings sheet
 * @param {String} key - The preference key
 * @param {any} defaultValue - Default value if preference doesn't exist
 * @return {any} The preference value
 */
function getUserPreference(key, defaultValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Settings sheet if it doesn't exist
  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.getRange("A1:B1").setValues([["Preference", "Value"]]);
    settingsSheet.getRange("A1:B1").setFontWeight("bold");
    // Hide the Settings sheet to keep the UI clean
    settingsSheet.hideSheet();
  }
  
  // Look for the preference key
  const data = settingsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  
  // If not found, return default value
  return defaultValue;
}

/**
 * Sets a user preference value in the Settings sheet
 * @param {String} key - The preference key
 * @param {any} value - The preference value to set
 */
function setUserPreference(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Settings sheet if it doesn't exist
  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.getRange("A1:B1").setValues([["Preference", "Value"]]);
    settingsSheet.getRange("A1:B1").setFontWeight("bold");
    // Hide the Settings sheet to keep the UI clean
    settingsSheet.hideSheet();
  }
  
  // Look for the preference key to update
  const data = settingsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      settingsSheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  
  // If key not found, append new row
  const lastRow = settingsSheet.getLastRow();
  settingsSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
}

/**
 * Toggles the showing of sub-categories in the Overview sheet
 */
function toggleShowSubCategories() {
  const currentValue = getUserPreference("ShowSubCategories", true);
  setUserPreference("ShowSubCategories", !currentValue);
  
  // Regenerate the overview with the new setting
  createFinancialOverview();
  
  // Show a toast message to confirm the change
  const status = !currentValue ? "enabled" : "disabled";
  SpreadsheetApp.getActiveSpreadsheet().toast(`Sub-categories ${status} in Overview sheet`, "Settings Updated");
}

/**
 * Sets budget targets for expense categories
 */
function setBudgetTargets() {
  // TODO: Implement budget targets setting functionality
  SpreadsheetApp.getUi().alert('Set Budget Targets - Coming Soon!');
}

function setupEmailReports() {
  // TODO: Implement email reports configuration
  SpreadsheetApp.getUi().alert('Setup Email Reports - Coming Soon!');
}

function refreshCache() {
  // Preserve the existing refresh cache functionality
  // TODO: Implement actual cache refresh logic
  SpreadsheetApp.getUi().alert('Refreshing Dropdown Cache - Coming Soon!');
}

// Functions are automatically global in Google Apps Script
