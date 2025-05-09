/**
 * @fileoverview UI Service Module for Financial Planning Tools.
 * Provides centralized UI functionality like notifications, dialogs, and sidebars.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const UIServiceModule = (function() {
  /**
   * Constructor for the UIServiceModule.
   * @constructor
   */
  function UIServiceModuleConstructor() {
    // No explicit dependencies needed for now based on current methods.
    // If config (e.g., for default titles, sizes) is needed later,
    // it can be injected here.
  }

  UIServiceModuleConstructor.prototype.showLoadingSpinner = function(message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Working...');
  };

  UIServiceModuleConstructor.prototype.hideLoadingSpinner = function() {
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
  };

  UIServiceModuleConstructor.prototype.showSuccessNotification = function(message, duration = 5) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Success', duration);
  };

  UIServiceModuleConstructor.prototype.showErrorNotification = function(title, message) {
    SpreadsheetApp.getUi().alert(`${title}: ${message}`);
  };

  UIServiceModuleConstructor.prototype.showInfoAlert = function(title, message) {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  };

  UIServiceModuleConstructor.prototype.showConfirmationDialog = function(title, message) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(title, message, ui.ButtonSet.OK_CANCEL);
    return response === ui.Button.OK;
  };

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

  UIServiceModuleConstructor.prototype.showSidebar = function(title, htmlContent) {
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setTitle(title)
      .setWidth(300); // Default width
    ui.showSidebar(htmlOutput);
  };

  UIServiceModuleConstructor.prototype.showModalDialog = function(title, htmlContent, width = 600, height = 400) {
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(width)
      .setHeight(height);
    ui.showModalDialog(htmlOutput, title);
  };

  return UIServiceModuleConstructor;
})();
