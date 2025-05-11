/**
 * @fileoverview Error Service Module for Financial Planning Tools.
 * Provides centralized error handling, logging, and reporting.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const ErrorServiceModule = (function() {
  /**
   * Custom error class for application-specific errors.
   * @class FinancialPlannerError
   * @extends Error
   */
  class FinancialPlannerError extends Error {
    constructor(message, details = {}) {
      super(message);
      this.name = 'FinancialPlannerError';
      this.details = details;
      this.timestamp = new Date();
    }
  }

  /**
   * Logs an error object to a designated Google Sheet.
   * @param {Error|FinancialPlannerError} error - The error object to log.
   * @param {object} configService - The Config service instance.
   * @private
   */
  function logToSheet(error, configService) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      // Assumes FinancialPlanner.Utils is available globally or will be refactored.
      // If Utils becomes a class, it might need to be injected as well.
      const errorSheet = FinancialPlanner.Utils.getOrCreateSheet(ss, configService.getSheetNames().ERROR_LOG);

      if (errorSheet.getLastRow() === 0) {
        errorSheet.appendRow(['Timestamp', 'Error Type', 'Message', 'Details']);
        errorSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      }

      const errorDetails = error.details || {};
      const formattedDetails = JSON.stringify(errorDetails);

      errorSheet.appendRow([
        error.timestamp || new Date(),
        error.name || 'Error',
        error.message,
        formattedDetails,
      ]);

      const lastRow = errorSheet.getLastRow();
      errorSheet.getRange(lastRow, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');

      const severity = errorDetails.severity || 'low';
      const bgColor = severity === 'high' ? '#F9BDBD' :
                      severity === 'medium' ? '#FFE0B2' : '#E1F5FE';
      errorSheet.getRange(lastRow, 1, 1, 4).setBackground(bgColor);
    } catch (logError) {
      console.error('Failed to log error to sheet:', logError);
      console.error('Original error:', error.message, error.details);
    }
  }

  /**
   * Logs an error object to the Google Apps Script console.
   * @param {Error|FinancialPlannerError} error - The error object to log.
   * @private
   */
  function logToConsole(error) {
    console.error(`[${error.name || 'Error'}] ${error.message}`);
    if (error.details) {
      console.error('Details:', error.details);
    }
    if (error.stack) {
      console.error('Stack trace:', error.stack);
    }
  }

  /**
   * Constructor for the ErrorServiceModule.
   * @param {object} configInstance - An instance of ConfigModule.
   * @param {object} uiServiceInstance - An instance of UIServiceModule.
   * @constructor
   */
  function ErrorServiceModuleConstructor(configInstance, uiServiceInstance) {
    this.config = configInstance;
    this.uiService = uiServiceInstance; // Will be used by this.handle
  }

  ErrorServiceModuleConstructor.prototype.create = function(message, details = {}) {
    return new FinancialPlannerError(message, details);
  };

  ErrorServiceModuleConstructor.prototype.log = function(error) {
    logToConsole(error);
    // Pass the config instance to logToSheet
    logToSheet(error, this.config);
  };

  ErrorServiceModuleConstructor.prototype.handle = function(error, userFriendlyMessage) {
    this.log(error);
    // Ensure uiService is available (it will be once UIServiceModule is refactored and injected)
    if (this.uiService && typeof this.uiService.showErrorNotification === 'function') {
      this.uiService.showErrorNotification(
        'Error',
        userFriendlyMessage || error.message
      );
    } else {
      console.error('UIService not available or showErrorNotification is not a function. Cannot show UI error.');
      // Fallback to a simple toast if possible, or just rely on console/sheet log
      SpreadsheetApp.getActiveSpreadsheet().toast(userFriendlyMessage || error.message, 'Error Occurred', 5);
    }
  };

  ErrorServiceModuleConstructor.prototype.wrap = function(fn, userFriendlyMessage) {
    const self = this;
    return function(...args) {
      try {
        return fn.apply(this, args);
      } catch (error) {
        self.handle(
          error,
          userFriendlyMessage || 'An error occurred while performing the operation.'
        );
        throw error;
      }
    };
  };
  
  // Expose FinancialPlannerError class if it needs to be used for `instanceof` checks externally
  // ErrorServiceModuleConstructor.FinancialPlannerError = FinancialPlannerError;


  return ErrorServiceModuleConstructor;
})();
