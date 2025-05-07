/**
 * Financial Planning Tools - Namespace Definition
 * 
 * This file defines the global namespace for the Financial Planning Tools project.
 * It helps prevent global namespace pollution by encapsulating all functionality
 * within a single global object.
 */

// Create a global namespace
var FinancialPlanner = FinancialPlanner || {};

// Add version information
FinancialPlanner.VERSION = '1.0.0';

// Add metadata
FinancialPlanner.META = {
  name: 'Financial Planning Tools',
  description: 'Google Apps Script project for financial planning and analysis',
  author: 'Financial Planning Team',
  lastUpdated: '2025-05-07'
};

/**
 * Utility function to safely extend the namespace with a new module
 * @param {String} moduleName - The name of the module to create
 * @param {Function} moduleFactory - Factory function that returns the module
 * @returns {Object} The created module
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
