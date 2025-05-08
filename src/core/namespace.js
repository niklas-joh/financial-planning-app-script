/**
 * Financial Planning Tools - Namespace Definition
 * 
 * This file defines the global namespace for the Financial Planning Tools project.
 * It helps prevent global namespace pollution by encapsulating all functionality
 * within a single global object.
 */

/**
 * The global namespace for the Financial Planning Tools project.
 * All modules, services, and utility functions should be attached to this object
 * to prevent polluting the global scope.
 * @namespace FinancialPlanner
 */
var FinancialPlanner = FinancialPlanner || {};

/**
 * The current version of the Financial Planning Tools application.
 * @memberof FinancialPlanner
 * @type {string}
 * @const
 */
FinancialPlanner.VERSION = '1.0.0';

/**
 * Metadata about the Financial Planning Tools application.
 * @memberof FinancialPlanner
 * @type {object}
 * @property {string} name - The official name of the application.
 * @property {string} description - A brief description of the application.
 * @property {string} author - The author or team responsible for the application.
 * @property {string} lastUpdated - The date of the last significant update (YYYY-MM-DD).
 * @const
 */
FinancialPlanner.META = {
  name: 'Financial Planning Tools',
  description: 'Google Apps Script project for financial planning and analysis',
  author: 'Financial Planning Team',
  lastUpdated: '2025-05-07'
};

/**
 * Utility function to safely create and attach a new module to the `FinancialPlanner` namespace.
 * This function takes a module name and a factory function. The factory function is executed,
 * and its return value (the module object) is assigned to `FinancialPlanner[moduleName]`.
 * A warning is logged if a module with the same name already exists.
 *
 * @memberof FinancialPlanner
 * @param {string} moduleName - The desired name for the new module (e.g., "Utils", "ReportService").
 * @param {function(): object} moduleFactory - A factory function that, when called, returns the module object.
 *                                           This typically involves an IIFE (Immediately Invoked Function Expression).
 * @return {object} The newly created and attached module.
 *
 * @example
 * // Creating a simple 'Logger' module
 * FinancialPlanner.createModule('Logger', function() {
 *   // Private stuff
 *   const logHistory = [];
 *
 *   // Public API
 *   return {
 *     log: function(message) {
 *       console.log(message);
 *       logHistory.push(message);
 *     },
 *     getHistory: function() {
 *       return logHistory;
 *     }
 *   };
 * });
 *
 * FinancialPlanner.Logger.log("Module created!");
 */
FinancialPlanner.createModule = function(moduleName, moduleFactory) {
  // Check if module already exists
  if (this[moduleName] !== undefined) {
    console.warn(`Module ${moduleName} already exists and will be overwritten.`);
  }
  
  // Create the module using the factory function
  this[moduleName] = moduleFactory();
  
  // Return the created module for chaining
  return this[moduleName];
};
