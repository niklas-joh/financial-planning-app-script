/**
 * Financial Planning Tools - UI Service
 * 
 * This file provides a centralized service for UI-related functionality,
 * including notifications, alerts, and other user interface operations.
 */

// Create the UIService module within the FinancialPlanner namespace
FinancialPlanner.UIService = (function() {
  // Private variables and functions can be defined here
  
  // Public API
  return {
    /**
     * Shows a loading spinner with a message
     * @param {String} message - Message to display
     */
    showLoadingSpinner: function(message) {
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Working...");
    },
    
    /**
     * Hides the loading spinner
     */
    hideLoadingSpinner: function() {
      // Google Apps Script doesn't have a direct way to dismiss toasts
      // So we just show a blank toast that will disappear quickly
      SpreadsheetApp.getActiveSpreadsheet().toast("", "", 1);
    },
    
    /**
     * Shows a success notification
     * @param {String} message - Success message
     * @param {Number} duration - Duration in seconds to show the message (default: 5)
     */
    showSuccessNotification: function(message, duration = 5) {
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Success", duration);
    },
    
    /**
     * Shows an error notification
     * @param {String} title - Error title
     * @param {String} message - Error message
     */
    showErrorNotification: function(title, message) {
      SpreadsheetApp.getUi().alert(`${title}: ${message}`);
    },
    
    /**
     * Shows an information alert
     * @param {String} title - Alert title
     * @param {String} message - Alert message
     */
    showInfoAlert: function(title, message) {
      SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
    },
    
    /**
     * Shows a confirmation dialog
     * @param {String} title - Dialog title
     * @param {String} message - Dialog message
     * @return {Boolean} True if the user clicked "OK", false otherwise
     */
    showConfirmationDialog: function(title, message) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(title, message, ui.ButtonSet.OK_CANCEL);
      return response === ui.Button.OK;
    },
    
    /**
     * Shows a prompt dialog to get user input
     * @param {String} title - Dialog title
     * @param {String} message - Dialog message
     * @param {String} defaultValue - Default value for the input field
     * @return {String|null} The user's input, or null if the user clicked "Cancel"
     */
    showPromptDialog: function(title, message, defaultValue = "") {
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL, defaultValue);
      
      if (response.getSelectedButton() === ui.Button.OK) {
        return response.getResponseText();
      }
      
      return null;
    },
    
    /**
     * Creates a sidebar with HTML content
     * @param {String} title - Sidebar title
     * @param {String} htmlContent - HTML content for the sidebar
     */
    showSidebar: function(title, htmlContent) {
      const ui = SpreadsheetApp.getUi();
      const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setTitle(title)
        .setWidth(300);
      
      ui.showSidebar(htmlOutput);
    },
    
    /**
     * Creates a modal dialog with HTML content
     * @param {String} title - Dialog title
     * @param {String} htmlContent - HTML content for the dialog
     * @param {Number} width - Dialog width in pixels (default: 600)
     * @param {Number} height - Dialog height in pixels (default: 400)
     */
    showModalDialog: function(title, htmlContent, width = 600, height = 400) {
      const ui = SpreadsheetApp.getUi();
      const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(width)
        .setHeight(height);
      
      ui.showModalDialog(htmlOutput, title);
    }
  };
})();
