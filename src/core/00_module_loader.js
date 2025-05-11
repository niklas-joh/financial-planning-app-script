/**
 * @fileoverview Initializes the FinancialPlanner namespace and orchestrates the loading
 * and instantiation of core application modules. This script ensures that modules
 * are loaded in the correct order and dependencies are injected appropriately.
 *
 * It is intended to be the first script loaded by Google Apps Script to prevent
 * issues related to script execution order and undefined namespaces or modules.
 */

// Ensure the global FinancialPlanner namespace exists.
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * The current version of the Financial Planning Tools application.
 * @memberof FinancialPlanner
 * @type {string}
 * @const
 */
FinancialPlanner.VERSION = '3.0.0';

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
  lastUpdated: '2025-05-11'
};

// Instantiate core modules in the correct dependency order.
FinancialPlanner.Config = new ConfigModule();
FinancialPlanner.UIService = new UIServiceModule();
FinancialPlanner.ErrorService = new ErrorServiceModule(FinancialPlanner.Config, FinancialPlanner.UIService);
FinancialPlanner.CacheService = new CacheServiceModule(FinancialPlanner.Config, FinancialPlanner.ErrorService);
FinancialPlanner.SettingsService = new SettingsServiceModule(FinancialPlanner.Config, FinancialPlanner.UIService, FinancialPlanner.ErrorService);

// Instantiate new services for refactoring
FinancialPlanner.FormulaBuilder = new FormulaBuilderModule(FinancialPlanner.Config);
FinancialPlanner.SheetBuilder = new SheetBuilderModule(FinancialPlanner.Config, FinancialPlanner.Utils);
FinancialPlanner.MetricsCalculator = new MetricsCalculatorModule(FinancialPlanner.Config);
FinancialPlanner.DataProcessor = new DataProcessorModule(FinancialPlanner.Config, FinancialPlanner.ErrorService);

// Instantiate controllers last as they depend on services
FinancialPlanner.Controllers = new ControllersModule(FinancialPlanner.Config, FinancialPlanner.UIService, FinancialPlanner.ErrorService);

// Log successful initialization
Logger.log(`FinancialPlanner modules loaded successfully. Version: ${FinancialPlanner.VERSION}`);
