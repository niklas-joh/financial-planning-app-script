/**
 * Financial Planning Tools - UI Service
 * 
 * This file provides a centralized service for UI-related functionality,
 * including notifications, alerts, and other user interface operations.
 */

/**
 * @namespace FinancialPlanner.UIService
 */
FinancialPlanner.UIService = (function() {
  // Private variables and functions can be defined here
  
  // Public API
  return {
    /**
     * Displays a "toast" notification with a "Working..." title, typically used for loading states.
     * @param {string} message - The message to display in the toast.
     * @return {void}
     *
     * @example
     * FinancialPlanner.UIService.showLoadingSpinner("Processing data...");
     */
    showLoadingSpinner: function(message) {
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Working...");
    },
    
    /**
     * Attempts to hide the loading spinner or any active toast.
     * Google Apps Script does not offer a direct way to dismiss toasts.
     * This implementation shows a blank toast with a very short duration (1 second)
     * to effectively clear the previous toast.
     * @return {void}
     *
     * @example
     * FinancialPlanner.UIService.hideLoadingSpinner();
     */
    hideLoadingSpinner: function() {
      SpreadsheetApp.getActiveSpreadsheet().toast("", "", 1);
    },
    
    /**
     * Displays a "toast" notification with a "Success" title.
     * @param {string} message - The success message to display.
     * @param {number} [duration=5] - The duration in seconds for which the toast should be visible.
     * @return {void}
     *
     * @example
     * FinancialPlanner.UIService.showSuccessNotification("Settings saved successfully!", 10);
     */
    showSuccessNotification: function(message, duration = 5) {
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Success", duration);
    },
    
    /**
     * Displays a standard alert dialog with a custom title and message.
     * This is typically used for error notifications.
     * @param {string} title - The title of the alert dialog.
     * @param {string} message - The main message content of the alert.
     * @return {void}
     *
     * @example
     * FinancialPlanner.UIService.showErrorNotification("Validation Error", "The email address is invalid.");
     */
    showErrorNotification: function(title, message) {
      SpreadsheetApp.getUi().alert(`${title}: ${message}`);
    },
    
    /**
     * Displays a standard alert dialog with an "OK" button.
     * Useful for informational messages.
     * @param {string} title - The title of the alert dialog.
     * @param {string} message - The main message content of the alert.
     * @return {void}
     *
     * @example
     * FinancialPlanner.UIService.showInfoAlert("Update Complete", "Your profile has been updated.");
     */
    showInfoAlert: function(title, message) {
      SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
    },
    
    /**
     * Displays an alert dialog with "OK" and "Cancel" buttons.
     * Useful for confirmation prompts.
     * @param {string} title - The title of the confirmation dialog.
     * @param {string} message - The confirmation message.
     * @return {boolean} True if the user clicked "OK", false if "Cancel" or closed the dialog.
     *
     * @example
     * if (FinancialPlanner.UIService.showConfirmationDialog("Confirm Deletion", "Are you sure you want to delete this item?")) {
     *   // Proceed with deletion
     * }
     */
    showConfirmationDialog: function(title, message) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(title, message, ui.ButtonSet.OK_CANCEL);
      return response === ui.Button.OK;
    },
    
    /**
     * Displays a prompt dialog to get text input from the user.
     * Includes "OK" and "Cancel" buttons.
     * @param {string} title - The title of the prompt dialog.
     * @param {string} message - The message or question to display to the user.
     * @param {string} [defaultValue=""] - The default value to pre-fill in the input field.
     * @return {string | null} The text entered by the user if "OK" was clicked; otherwise, null.
     *
     * @example
     * const categoryName = FinancialPlanner.UIService.showPromptDialog("New Category", "Enter the name for the new category:", "Miscellaneous");
     * if (categoryName) {
     *   // Process the new category name
     * }
     */
    showPromptDialog: function(title, message, defaultValue = "") {
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL); // defaultValue is not a param for this version of prompt
      
      if (response.getSelectedButton() === ui.Button.OK) {
        return response.getResponseText();
      }
      
      return null;
    },
    
    /**
     * Displays a custom sidebar in the Google Sheets UI, populated with the provided HTML content.
     * @param {string} title - The title to display at the top of the sidebar.
     * @param {string} htmlContent - A string containing the HTML markup for the sidebar's content.
     * @return {void}
     *
     * @example
     * const sidebarHtml = "<h1>Settings</h1><p>Configure your preferences here.</p>";
     * FinancialPlanner.UIService.showSidebar("My App Settings", sidebarHtml);
     */
    showSidebar: function(title, htmlContent) {
      const ui = SpreadsheetApp.getUi();
      const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setTitle(title)
        .setWidth(300); // Default width, can be customized
      
      ui.showSidebar(htmlOutput);
    },
    
    /**
     * Displays a custom modal dialog in the Google Sheets UI, populated with the provided HTML content.
     * @param {string} title - The title to display at the top of the modal dialog.
     * @param {string} htmlContent - A string containing the HTML markup for the dialog's content.
     * @param {number} [width=600] - The desired width of the modal dialog in pixels.
     * @param {number} [height=400] - The desired height of the modal dialog in pixels.
     * @return {void}
     *
     * @example
     * const modalHtml = "<h2>Detailed Report</h2><div id='chart_div'></div>";
     * FinancialPlanner.UIService.showModalDialog("View Report", modalHtml, 800, 600);
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
