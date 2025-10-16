/**
 * @fileoverview Error Service Module for Financial Planning Tools.
 * Provides centralized error handling, logging (to console and a designated sheet),
 * and user-friendly error reporting. It includes a custom error class
 * `FinancialPlannerError` for application-specific errors.
 * @module services/error-service
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Error Service - Provides centralized error handling and logging.
 * Uses IIFE to keep FinancialPlannerError class and helper functions private.
 * @namespace FinancialPlanner.ErrorService
 */
FinancialPlanner.ErrorService = (function() {
  /**
   * Custom error class for application-specific errors within Financial Planning Tools.
   * Extends the native `Error` class to include additional details and a timestamp.
   * @class FinancialPlannerError
   * @extends Error
   * @param {string} message - The human-readable description of the error.
   * @param {object} [details={}] - An optional object containing additional contextual
   *   information about the error (e.g., severity, originalError, relevant data).
   * @private
   */
  class FinancialPlannerError extends Error {
    constructor(message, details) {
      details = details || {};
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
   * @private
   */
  function logToSheet(error) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const errorSheet = FinancialPlanner.Utils.getOrCreateSheet(
        ss, 
        FinancialPlanner.Config.getSheetNames().ERROR_LOG
      );

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
    console.error('[' + (error.name || 'Error') + '] ' + error.message);
    if (error.details) {
      console.error('Details:', error.details);
    }
    if (error.stack) {
      console.error('Stack trace:', error.stack);
    }
  }

  // Public API
  return {
    /**
     * Creates a new `FinancialPlannerError` instance.
     * This is the preferred way to generate application-specific errors.
     * @param {string} message - The human-readable error message.
     * @param {object} [details={}] - Optional. An object containing additional details
     *   (e.g., severity: 'low'|'medium'|'high', originalError, context).
     * @returns {FinancialPlannerError} A new instance of `FinancialPlannerError`.
     * @memberof FinancialPlanner.ErrorService
     */
    create: function(message, details) {
      details = details || {};
      return new FinancialPlannerError(message, details);
    },

    /**
     * Logs an error to both the console and the designated error log sheet.
     * @param {Error|FinancialPlannerError} error - The error object to log.
     * @memberof FinancialPlanner.ErrorService
     */
    log: function(error) {
      logToConsole(error);
      logToSheet(error);
    },

    /**
     * Handles an error by logging it and displaying a user-friendly message
     * via the UIService. If UIService is unavailable, falls back to a simple toast.
     * @param {Error|FinancialPlannerError} error - The error object to handle.
     * @param {string} [userFriendlyMessage] - An optional user-friendly message to display.
     *   If not provided, the error's message property is used.
     * @memberof FinancialPlanner.ErrorService
     */
    handle: function(error, userFriendlyMessage) {
      this.log(error);
      
      // Access UIService via namespace
      if (FinancialPlanner.UIService && typeof FinancialPlanner.UIService.showErrorNotification === 'function') {
        FinancialPlanner.UIService.showErrorNotification(
          'Error',
          userFriendlyMessage || error.message
        );
      } else {
        console.error('UIService not available or showErrorNotification is not a function. Cannot show UI error.');
        // Fallback to a simple toast if possible, or just rely on console/sheet log
        SpreadsheetApp.getActiveSpreadsheet().toast(userFriendlyMessage || error.message, 'Error Occurred', 5);
      }
    },

    /**
     * Wraps a function with error handling. If the wrapped function throws an error,
     * this handler will catch it, log it using `this.handle()`, and then re-throw the error.
     * This is useful for ensuring consistent error handling around functions that might fail.
     * @param {function(...*): *} fn - The function to wrap with error handling.
     * @param {string} [userFriendlyMessage] - An optional user-friendly message to display
     *   if the wrapped function throws an error. Defaults to a generic message.
     * @returns {function(...*): *} The wrapped function, which includes error handling.
     * @memberof FinancialPlanner.ErrorService
     */
    wrap: function(fn, userFriendlyMessage) {
      const self = this;
      return function() {
        try {
          return fn.apply(this, arguments);
        } catch (error) {
          self.handle(
            error,
            userFriendlyMessage || 'An error occurred while performing the operation.'
          );
          throw error; // Re-throw the error after handling, so callers can react if needed.
        }
      };
    }
  };
})();
