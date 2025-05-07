/**
 * Financial Planning Tools - Error Service
 * 
 * This file provides a centralized service for error handling, logging, and reporting.
 * It helps standardize error handling across the application.
 */

// Create the ErrorService module within the FinancialPlanner namespace
FinancialPlanner.ErrorService = (function(config, uiService) {
  // Private variables and functions
  
  /**
   * Custom error class for Financial Planning Tools
   * @class
   * @extends Error
   */
  class FinancialPlannerError extends Error {
    /**
     * Creates a new FinancialPlannerError
     * @param {String} message - Error message
     * @param {Object} details - Additional error details
     */
    constructor(message, details = {}) {
      super(message);
      this.name = 'FinancialPlannerError';
      this.details = details;
      this.timestamp = new Date();
    }
  }
  
  /**
   * Logs an error to the error log sheet
   * @param {Error} error - The error to log
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
   * Logs an error to the console
   * @param {Error} error - The error to log
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
     * Creates a new FinancialPlannerError
     * @param {String} message - Error message
     * @param {Object} details - Additional error details
     * @return {FinancialPlannerError} The created error
     */
    create: function(message, details = {}) {
      return new FinancialPlannerError(message, details);
    },
    
    /**
     * Logs an error to both the error log sheet and console
     * @param {Error} error - The error to log
     */
    log: function(error) {
      // Log to console first (this will always work)
      logToConsole(error);
      
      // Then try to log to sheet
      logToSheet(error);
    },
    
    /**
     * Handles an error by logging it and showing a notification to the user
     * @param {Error} error - The error to handle
     * @param {String} userFriendlyMessage - A user-friendly message to display
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
     * Wraps a function with error handling
     * @param {Function} fn - The function to wrap
     * @param {String} userFriendlyMessage - A user-friendly message to display if an error occurs
     * @return {Function} The wrapped function
     */
    wrap: function(fn, userFriendlyMessage) {
      return function() {
        try {
          return fn.apply(this, arguments);
        } catch (error) {
          FinancialPlanner.ErrorService.handle(
            error,
            userFriendlyMessage || "An error occurred while performing the operation."
          );
          throw error; // Re-throw to allow caller to handle if needed
        }
      };
    }
  };
})(FinancialPlanner.Config, FinancialPlanner.UIService);
