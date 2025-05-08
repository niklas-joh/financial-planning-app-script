/**
 * Financial Planning Tools - Error Service
 * 
 * This file provides a centralized service for error handling, logging, and reporting.
 * It helps standardize error handling across the application.
 */

// Create the ErrorService module within the FinancialPlanner namespace
/**
 * @namespace FinancialPlanner.ErrorService
 * @param {FinancialPlanner.Config} config - The configuration service, used for getting sheet names (e.g., error log sheet).
 * @param {FinancialPlanner.UIService} uiService - The UI service, used for displaying error notifications to the user.
 */
FinancialPlanner.ErrorService = (function(config, uiService) {
  // Private variables and functions
  
  /**
   * Custom error class for application-specific errors within Financial Planning Tools.
   * Extends the native `Error` class to include additional details and a timestamp.
   * @class FinancialPlannerError
   * @extends Error
   */
  class FinancialPlannerError extends Error {
    /**
     * Creates an instance of FinancialPlannerError.
     * @param {string} message - The primary error message.
     * @param {object} [details={}] - An optional object containing additional details about the error (e.g., severity, context).
     */
    constructor(message, details = {}) {
      super(message);
      this.name = 'FinancialPlannerError';
      this.details = details;
      this.timestamp = new Date();
    }
  }
  
  /**
   * Logs an error object to a designated Google Sheet (defined in `config.getSheetNames().ERROR_LOG`).
   * If the sheet doesn't exist, it's created. Includes timestamp, error type, message, and details.
   * Errors during logging to the sheet are caught and logged to the console.
   * @param {Error|FinancialPlannerError} error - The error object to log.
   * @private
   */
  function logToSheet(error) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const errorSheet = FinancialPlanner.Utils.getOrCreateSheet(ss, config.getSheetNames().ERROR_LOG);
      
      // Create headers if this is a new sheet
      if (errorSheet.getLastRow() === 0) {
        errorSheet.appendRow(["Timestamp", "Error Type", "Message", "Details"]);
        errorSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
      }
      
      // Format error details for logging
      const errorDetails = error.details || {};
      const formattedDetails = JSON.stringify(errorDetails);
      
      // Append error information
      errorSheet.appendRow([
        error.timestamp || new Date(), 
        error.name || "Error", 
        error.message, 
        formattedDetails
      ]);
      
      // Format the timestamp
      const lastRow = errorSheet.getLastRow();
      errorSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
      
      // Set colors based on error severity
      const severity = errorDetails.severity || "low";
      const bgColor = severity === "high" ? "#F9BDBD" : 
                      severity === "medium" ? "#FFE0B2" : "#E1F5FE";
      errorSheet.getRange(lastRow, 1, 1, 4).setBackground(bgColor);
    } catch (logError) {
      // If we can't log to sheet, at least log to console
      console.error("Failed to log error to sheet:", logError);
      console.error("Original error:", error.message, error.details);
    }
  }
  
  /**
   * Logs an error object to the Google Apps Script console.
   * Includes the error name, message, details (if any), and stack trace (if available).
   * @param {Error|FinancialPlannerError} error - The error object to log.
   * @private
   */
  function logToConsole(error) {
    console.error(`[${error.name || "Error"}] ${error.message}`);
    
    if (error.details) {
      console.error("Details:", error.details);
    }
    
    if (error.stack) {
      console.error("Stack trace:", error.stack);
    }
  }
  
  // Public API
  return {
    /**
     * Creates a new instance of `FinancialPlannerError`.
     * @param {string} message - The primary error message.
     * @param {object} [details={}] - An optional object containing additional details about the error.
     * @return {FinancialPlannerError} The newly created `FinancialPlannerError` object.
     *
     * @example
     * const customError = FinancialPlanner.ErrorService.create("Configuration not found", { setting: "API_KEY" });
     * throw customError;
     */
    create: function(message, details = {}) {
      return new FinancialPlannerError(message, details);
    },
    
    /**
     * Logs an error to both the Google Apps Script console and the designated error log sheet.
     * @param {Error|FinancialPlannerError} error - The error object to log.
     *
     * @example
     * try {
     *   // some operation
     * } catch (e) {
     *   FinancialPlanner.ErrorService.log(e);
     * }
     */
    log: function(error) {
      // Log to console first (this will always work)
      logToConsole(error);
      
      // Then try to log to sheet
      logToSheet(error);
    },
    
    /**
     * Handles an error by logging it (using `this.log()`) and then displaying a user-friendly
     * notification via `uiService.showErrorNotification()`.
     * @param {Error|FinancialPlannerError} error - The error object to handle.
     * @param {string} [userFriendlyMessage] - An optional user-friendly message to display.
     *                                         If not provided, `error.message` is used.
     *
     * @example
     * try {
     *   // some operation
     * } catch (e) {
     *   FinancialPlanner.ErrorService.handle(e, "Sorry, something went wrong while processing your request.");
     * }
     */
    handle: function(error, userFriendlyMessage) {
      // Log the error
      this.log(error);
      
      // Show a notification to the user
      uiService.showErrorNotification(
        "Error",
        userFriendlyMessage || error.message
      );
    },
    
    /**
     * Wraps a given function with a try-catch block. If the wrapped function throws an error,
     * it's caught, handled by `FinancialPlanner.ErrorService.handle()`, and then re-thrown.
     * @param {function(...any): any} fn - The function to wrap with error handling.
     * @param {string} [userFriendlyMessage] - An optional user-friendly message to display if an error occurs.
     *                                         Defaults to "An error occurred while performing the operation.".
     * @return {function(...any): any} The wrapped function.
     * @throws {Error} Re-throws the original error after handling.
     *
     * @example
     * const safeFunction = FinancialPlanner.ErrorService.wrap(function() {
     *   // Potentially risky code
     *   if (Math.random() < 0.5) throw new Error("Random failure!");
     *   return "Success!";
     * }, "Operation failed. Please try again.");
     *
     * try {
     *   safeFunction();
     * } catch (e) {
     *   // Error already handled (logged and UI notification shown),
     *   // but can still perform additional cleanup if needed.
     *   console.log("Caught re-thrown error in caller.");
     * }
     */
    wrap: function(fn, userFriendlyMessage) {
      const self = this; // Capture 'this' context for ErrorService.handle
      return function() {
        try {
          return fn.apply(this, arguments);
        } catch (error) {
          self.handle( // Use captured 'self'
            error,
            userFriendlyMessage || "An error occurred while performing the operation."
          );
          throw error; // Re-throw to allow caller to handle if needed
        }
      };
    }
  };
})(FinancialPlanner.Config, FinancialPlanner.UIService);
