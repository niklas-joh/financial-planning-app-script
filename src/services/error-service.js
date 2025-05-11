/**
 * @fileoverview Error Service Module for Financial Planning Tools.
 * Provides centralized error handling, logging (to console and a designated sheet),
 * and user-friendly error reporting. It includes a custom error class
 * `FinancialPlannerError` for application-specific errors.
 * This module is designed to be instantiated by `00_module_loader.js`.
 * @module services/error-service
 */

/**
 * IIFE to encapsulate the ErrorServiceModule logic.
 * @returns {function} The ErrorServiceModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const ErrorServiceModule = (function() {
  /**
   * Custom error class for application-specific errors within Financial Planning Tools.
   * Extends the native `Error` class to include additional details and a timestamp.
   * @class FinancialPlannerError
   * @extends Error
   * @param {string} message - The human-readable description of the error.
   * @param {object} [details={}] - An optional object containing additional contextual
   *   information about the error (e.g., severity, originalError, relevant data).
   */
  class FinancialPlannerError extends Error {
    constructor(message, details = {}) {
      super(message);
      /**
       * The name of the error type.
       * @type {string}
       * @default 'FinancialPlannerError'
       */
      this.name = 'FinancialPlannerError';
      /**
       * Additional details about the error.
       * @type {object}
       */
      this.details = details;
      /**
       * Timestamp of when the error occurred.
       * @type {Date}
       */
      this.timestamp = new Date();
    }
  }

  /**
   * Logs an error object to a designated Google Sheet specified in the configuration.
   * Creates the sheet and header row if they don't exist.
   * Formats the log entry and applies background color based on error severity.
   * @param {Error|FinancialPlannerError} error - The error object to log.
   * @param {ConfigModule} configService - The Config service instance, used to get sheet names.
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
   * Logs an error object to the Google Apps Script console (Logger or console.error).
   * Includes error name, message, details, and stack trace if available.
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
   * Initializes the service with configuration and UI service instances.
   * @param {ConfigModule} configInstance - An instance of ConfigModule.
   * @param {UIServiceModule} uiServiceInstance - An instance of UIServiceModule.
   * @constructor
   * @alias ErrorServiceModule
   * @memberof module:services/error-service
   */
  function ErrorServiceModuleConstructor(configInstance, uiServiceInstance) {
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
    this.uiService = uiServiceInstance; // Will be used by this.handle
  }

  /**
   * Creates a new `FinancialPlannerError` instance.
   * This is the preferred way to generate application-specific errors.
   * @param {string} message - The human-readable error message.
   * @param {object} [details={}] - Optional. An object containing additional details
   *   (e.g., severity: 'low'|'medium'|'high', originalError, context).
   * @returns {FinancialPlannerError} A new instance of `FinancialPlannerError`.
   * @memberof ErrorServiceModule
   */
  ErrorServiceModuleConstructor.prototype.create = function(message, details = {}) {
    return new FinancialPlannerError(message, details);
  };

  /**
   * Logs an error to both the console and the designated error log sheet.
   * @param {Error|FinancialPlannerError} error - The error object to log.
   * @memberof ErrorServiceModule
   */
  ErrorServiceModuleConstructor.prototype.log = function(error) {
    logToConsole(error);
    // Pass the config instance to logToSheet
    logToSheet(error, this.config);
  };

  /**
   * Handles an error by logging it and displaying a user-friendly message
   * via the UIService. If UIService is unavailable, falls back to a simple toast.
   * @param {Error|FinancialPlannerError} error - The error object to handle.
   * @param {string} [userFriendlyMessage] - An optional user-friendly message to display.
   *   If not provided, the error's message property is used.
   * @memberof ErrorServiceModule
   */
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

  /**
   * Wraps a function with error handling. If the wrapped function throws an error,
   * this handler will catch it, log it using `this.handle()`, and then re-throw the error.
   * This is useful for ensuring consistent error handling around functions that might fail.
   * @param {function(...*): *} fn - The function to wrap with error handling.
   * @param {string} [userFriendlyMessage] - An optional user-friendly message to display
   *   if the wrapped function throws an error. Defaults to a generic message.
   * @returns {function(...*): *} The wrapped function, which includes error handling.
   * @memberof ErrorServiceModule
   */
  ErrorServiceModuleConstructor.prototype.wrap = function(fn, userFriendlyMessage) {
    const self = this; // eslint-disable-line consistent-this
    return function(...args) {
      try {
        return fn.apply(this, args);
      } catch (error) {
        self.handle(
          error,
          userFriendlyMessage || 'An error occurred while performing the operation.'
        );
        throw error; // Re-throw the error after handling, so callers can react if needed.
      }
    };
  };
  
  // Expose FinancialPlannerError class if it needs to be used for `instanceof` checks externally
  // ErrorServiceModuleConstructor.FinancialPlannerError = FinancialPlannerError;
  // If FinancialPlannerError is to be exposed, it should be documented as part of the module.
  // For now, it's treated as an internal class.

  return ErrorServiceModuleConstructor;
})();
