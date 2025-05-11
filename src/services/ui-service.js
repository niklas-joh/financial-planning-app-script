/**
 * @fileoverview UI Service Module for Financial Planning Tools.
 * Provides a centralized interface for common UI interactions within Google Sheets,
 * such as displaying notifications, alerts, dialogs, sidebars, and spinners.
 * This module is designed to be instantiated by `00_module_loader.js`.
 * @module services/ui-service
 */

/**
 * IIFE to encapsulate the UIServiceModule logic.
 * @returns {function} The UIServiceModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const UIServiceModule = (function() {
  /**
   * Constructor for the UIServiceModule.
   * This service wraps Google Apps Script's `SpreadsheetApp.getUi()` and `HtmlService`
   * to provide a consistent API for UI elements.
   * @constructor
   * @alias UIServiceModule
   * @memberof module:services/ui-service
   */
  function UIServiceModuleConstructor() {
    // No explicit dependencies needed for now based on current methods.
    // If config (e.g., for default titles, sizes) is needed later,
    // it can be injected here.
  }

  /**
   * Displays a loading spinner (toast notification) to indicate an ongoing process.
   * @param {string} message - The message to display alongside the spinner (e.g., "Loading data...").
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.showLoadingSpinner = function(message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Working...');
  };

  /**
   * Hides any active loading spinner (toast notification).
   * Achieved by showing an empty toast for a very short duration.
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.hideLoadingSpinner = function() {
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Shows an empty toast for 1 second to clear previous
  };

  /**
   * Displays a success notification toast.
   * @param {string} message - The success message to display.
   * @param {number} [duration=5] - The duration in seconds for the toast to be visible.
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.showSuccessNotification = function(message, duration = 5) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Success', duration);
  };

  /**
   * Displays an error notification using a standard alert dialog.
   * @param {string} title - The title for the error dialog.
   * @param {string} message - The error message content.
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.showErrorNotification = function(title, message) {
    SpreadsheetApp.getUi().alert(`${title}: ${message}`);
  };

  /**
   * Displays an informational alert dialog with an OK button.
   * @param {string} title - The title for the alert dialog.
   * @param {string} message - The informational message content.
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.showInfoAlert = function(title, message) {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  };

  /**
   * Displays a confirmation dialog with OK and Cancel buttons.
   * @param {string} title - The title for the confirmation dialog.
   * @param {string} message - The message prompting for confirmation.
   * @returns {boolean} True if the user clicks OK, false if Cancel or dialog is closed.
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.showConfirmationDialog = function(title, message) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(title, message, ui.ButtonSet.OK_CANCEL);
    return response === ui.Button.OK;
  };

  /**
   * Displays a prompt dialog asking the user for input.
   * Note: The standard Google Apps Script `ui.prompt` does not support pre-filling the input field
   * with `defaultValue`. This parameter is kept for API consistency but is not used to pre-fill.
   * For pre-filled prompts, a custom HTML dialog is required.
   * @param {string} title - The title for the prompt dialog.
   * @param {string} message - The message/question to display to the user.
   * @param {string} [defaultValue=''] - Intended default value for the input (currently not used by `ui.prompt` for pre-fill).
   * @returns {string|null} The text entered by the user, or null if the user cancels or closes the dialog.
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.showPromptDialog = function(title, message, defaultValue = '') {
    const ui = SpreadsheetApp.getUi();
    // Note: The Apps Script `ui.prompt` method signature used here is:
    // prompt(title, prompt, buttons)
    // It does not directly take a defaultValue. If a pre-filled input is needed,
    // a custom HTML dialog (showModalDialog or showSidebar) would be more appropriate.
    // For simplicity, keeping the existing behavior which doesn't use defaultValue in the actual prompt.
    const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() === ui.Button.OK) {
      return response.getResponseText();
    }
    return null;
  };

  /**
   * Displays a custom sidebar in the Google Sheet interface.
   * @param {string} title - The title for the sidebar.
   * @param {string} htmlContent - The HTML string content to be displayed in the sidebar.
   *   This content is typically generated from an HTML file using `HtmlService.createHtmlOutputFromFile(fileName).getContent()`.
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.showSidebar = function(title, htmlContent) {
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setTitle(title)
      .setWidth(300); // Default width
    ui.showSidebar(htmlOutput);
  };

  /**
   * Displays a custom modal dialog in the Google Sheet interface.
   * @param {string} title - The title for the modal dialog.
   * @param {string} htmlContent - The HTML string content to be displayed in the dialog.
   *   Typically generated from an HTML file.
   * @param {number} [width=600] - The desired width of the dialog in pixels.
   * @param {number} [height=400] - The desired height of the dialog in pixels.
   * @memberof UIServiceModule
   */
  UIServiceModuleConstructor.prototype.showModalDialog = function(title, htmlContent, width = 600, height = 400) {
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(width)
      .setHeight(height);
    ui.showModalDialog(htmlOutput, title);
  };

  return UIServiceModuleConstructor;
})();
