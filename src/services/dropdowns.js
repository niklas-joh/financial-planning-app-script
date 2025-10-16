/**
 * Financial Planning Tools - Dropdown Service
 *
 * This file provides dynamic dependent dropdown functionality for the Transactions sheet.
 * It follows the namespace pattern and uses dependency injection for better maintainability.
 * @module services/dropdowns
 */

/**
 * @namespace FinancialPlanner.DropdownService
 * @description Service for managing dynamic dependent dropdowns in the 'Transactions' sheet.
 * It reads dropdown configurations from a 'Dropdowns' sheet, caches them, and applies
 * data validation rules to relevant cells based on user edits.
 * This service is designed as an IIFE and is attached to the `FinancialPlanner` global namespace.
 * @param {object} utils - The utility service, expected to be `FinancialPlanner.Utils`.
 * @param {object} uiService - The UI service for notifications, expected to be `FinancialPlanner.UIService`.
 * @param {object} errorService - The error handling service, expected to be `FinancialPlanner.ErrorService`.
 * @param {object} config - The global configuration service, expected to be `FinancialPlanner.Config`.
 */
FinancialPlanner.DropdownService = (function(utils, uiService, errorService, config) {
  /**
   * Configuration constants for the DropdownService.
   * @private
   * @readonly
   * @const {object} DROPDOWN_CONFIG
   * @property {number} CACHE_EXPIRY_SECONDS - Expiry time for the script cache.
   * @property {string} CACHE_KEY - Key used for storing dropdown data in script cache.
   * @property {object} SHEETS - Names of relevant sheets.
   * @property {string} SHEETS.TRANSACTIONS - Name of the sheet where dropdowns are active.
   * @property {string} SHEETS.DROPDOWNS - Name of the sheet containing dropdown definitions.
   * @property {object} COLUMNS - Column numbers (1-based) for Type, Category, and Sub-Category.
   * @property {number} COLUMNS.TYPE - Column for 'Type' dropdown.
   * @property {number} COLUMNS.CATEGORY - Column for 'Category' dropdown.
   * @property {number} COLUMNS.SUB_CATEGORY - Column for 'Sub-Category' dropdown.
   * @property {object} UI - UI related constants.
   * @property {string} UI.PLACEHOLDER_TEXT - Placeholder text for dropdowns.
   * @property {string} UI.PENDING_BACKGROUND - Background color for cells pending selection.
   * @property {string} KEY_SEPARATOR - Separator used for creating composite keys for caching.
   */
  const DROPDOWN_CONFIG = {
    CACHE_EXPIRY_SECONDS: 300, // Cache expires in 5 minutes
    CACHE_KEY: 'dropdownsData',
    SHEETS: {
      TRANSACTIONS: 'Transactions', // Ensure this aligns with global config or is passed if different
      DROPDOWNS: 'Dropdowns'       // Ensure this aligns with global config or is passed if different
    },
    COLUMNS: {
      TYPE: 3,        // Column C
      CATEGORY: 4,    // Column D
      SUB_CATEGORY: 5 // Column E
    },
    UI: {
      PLACEHOLDER_TEXT: 'Please select',
      PENDING_BACKGROUND: '#eeeeee'
    },
    KEY_SEPARATOR: '___' // Separator for composite keys
  };

  // Private cache (in-memory for the current session, script cache for longer persistence)
  /**
   * In-memory cache for the dropdown data structure.
   * Loaded from ScriptCache or built from the 'Dropdowns' sheet.
   * Structure: `{ typeToCategories: object, typeCategoryToSubCategories: object }`
   * @private
   * @type {null|{typeToCategories: object, typeCategoryToSubCategories: object}}
   */
  let dropdownCache = null;

  /**
   * Converts Set values within an object to Arrays. This is used when preparing
   * data from `buildDropdownCache` for JSON stringification and storage in ScriptCache.
   * @param {Object<string, Set<string>>} obj - The input object where property values are Sets.
   * @return {Object<string, Array<string>>} A new object where Set values have been converted to Arrays.
   * @private
   */
  function mapSetsToArrays(obj) {
    const out = {};
    for (let key in obj) {
      if (obj.hasOwnProperty(key)) {
        out[key] = Array.from(obj[key]);
      }
    }
    return out;
  }

  /**
   * Builds or retrieves the dropdown cache.
   * It first attempts to load from Google Apps Script's `CacheService`. If not found or expired,
   * it reads data from the 'Dropdowns' sheet, processes it into a structured format
   * (Type -> Categories, and Type+Category -> SubCategories), and stores it in the ScriptCache.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet object
   *   from which to read the 'Dropdowns' sheet if the cache is not populated.
   * @return {{typeToCategories: Object<string, Array<string>>, typeCategoryToSubCategories: Object<string, Array<string>>}}
   *   An object containing the dropdown mappings. Returns empty mappings if an error occurs.
   * @private
   */
  function buildDropdownCache(spreadsheet) {
    try {
      const scriptCache = CacheService.getScriptCache(); // Native Google Apps Script CacheService
      let cachedData = scriptCache.get(DROPDOWN_CONFIG.CACHE_KEY);

      if (cachedData) {
        Logger.log('DropdownService: Using cached dropdown data from ScriptCache.');
        const parsed = JSON.parse(cachedData);
        return {
          typeToCategories: parsed.typeToCategories || {},
          typeCategoryToSubCategories: parsed.typeCategoryToSubCategories || {},
        };
      }

      Logger.log('DropdownService: Building dropdown cache from sheet.');
      const dropdownsSheetName = (config && config.get && config.get().SHEETS && config.get().SHEETS.DROPDOWNS)
                                  ? config.get().SHEETS.DROPDOWNS
                                  : DROPDOWN_CONFIG.SHEETS.DROPDOWNS;
      const dropdownsSheet = spreadsheet.getSheetByName(dropdownsSheetName);

      if (!dropdownsSheet) {
        throw new Error(`Dropdowns definition sheet "${dropdownsSheetName}" not found.`);
      }

      const dropdownsData = dropdownsSheet.getDataRange().getValues();
      // Assuming header is 'Type', 'Category', 'Sub-Category'
      const headerRow = (dropdownsData.length > 0 && dropdownsData[0][0] === 'Type') ? 1 : 0;

      const typeToCategories = {};
      const typeCategoryToSubCategories = {};

      for (let i = headerRow; i < dropdownsData.length; i++) {
        const [type, category, subCategory] = dropdownsData[i];
        if (!type || !category) continue; // Skip if essential data is missing

        if (!typeToCategories[type]) typeToCategories[type] = new Set();
        typeToCategories[type].add(category);

        if (subCategory) {
          const key = `${type}${DROPDOWN_CONFIG.KEY_SEPARATOR}${category}`;
          if (!typeCategoryToSubCategories[key]) typeCategoryToSubCategories[key] = new Set();
          typeCategoryToSubCategories[key].add(subCategory);
        }
      }

      const cacheToStore = {
        typeToCategories: mapSetsToArrays(typeToCategories),
        typeCategoryToSubCategories: mapSetsToArrays(typeCategoryToSubCategories),
      };

      try {
        scriptCache.put(DROPDOWN_CONFIG.CACHE_KEY, JSON.stringify(cacheToStore), DROPDOWN_CONFIG.CACHE_EXPIRY_SECONDS);
        Logger.log('DropdownService: Dropdown data cached in ScriptCache.');
      } catch (e) {
        Logger.log('DropdownService: Failed to cache dropdown data in ScriptCache: ' + e.toString());
      }
      return cacheToStore;
    } catch (error) {
      Logger.log('DropdownService: Error building dropdown cache: ' + error.toString());
      errorService.handle(errorService.create('Error building dropdown cache', { originalError: error.toString() }), "Could not load dropdown options.");
      return { typeToCategories: {}, typeCategoryToSubCategories: {} }; // Fallback
    }
  }

  /**
   * Sets a data validation rule (dropdown list) for a given cell range.
   * The dropdown list will include `DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT` as the first option.
   * If `options` is empty or null, any existing validation on the range is cleared.
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range to apply the data validation to.
   * @param {Array<string>} options - An array of strings representing the dropdown options.
   * @private
   */
  function setDropdownValidation(range, options) {
    range.clearDataValidations();
    if (!options || options.length === 0) {
      return;
    }
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT, ...options], true)
      .setAllowInvalid(true) // Allow users to enter values not in the list
      .build();
    range.setDataValidation(rule);
  }

  /**
   * Clears the content of a cell if its current value is the `DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT`.
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range to check and potentially clear.
   * @private
   */
  function clearPlaceholderIfNeeded(range) {
    if (range.getValue() === DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT) {
      range.clearContent();
    }
  }

  /**
   * Applies a background highlight (using `DROPDOWN_CONFIG.UI.PENDING_BACKGROUND`) to a cell,
   * typically to indicate that a selection is pending or required.
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range to highlight.
   * @private
   */
  function highlightPending(range) {
    range.setBackground(DROPDOWN_CONFIG.UI.PENDING_BACKGROUND);
  }

  /**
   * Removes any background highlighting from a cell by setting its background to null.
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range from which to clear highlighting.
   * @private
   */
  function clearHighlight(range) {
    range.setBackground(null);
  }

  /**
   * Handles edits made to the 'Type' column in the 'Transactions' sheet.
   * It updates the 'Category' dropdown options based on the newly selected 'Type',
   * clears dependent 'Sub-Category' if 'Category' becomes invalid, and manages UI highlights.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The 'Transactions' sheet object.
   * @param {number} row - The 1-based row number of the cell that was edited.
   * @param {Object<string, Array<string>>} typeToCategories - The mapping of types to their available categories,
   *   obtained from the dropdown cache.
   * @private
   */
  function handleTypeEdit(sheet, row, typeToCategories) {
    const typeCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.TYPE);
    const categoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.CATEGORY);
    const subCategoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.SUB_CATEGORY);

    const type = typeCell.getValue();
    const oldCategoryValue = categoryCell.getValue();
    const categories = typeToCategories[type] || [];

    setDropdownValidation(categoryCell, categories);

    if (oldCategoryValue && (categories.length === 0 || !categories.includes(oldCategoryValue))) {
      categoryCell.clearContent();
      subCategoryCell.clearContent().clearDataValidations();
      highlightPending(categoryCell);
    } else if (categories.length > 0 && !oldCategoryValue) {
        highlightPending(categoryCell); // Highlight if new options available and cell is empty
    }


    clearPlaceholderIfNeeded(typeCell);
    if (categoryCell.getValue() !== DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT && categoryCell.getValue() !== "") {
        clearHighlight(categoryCell);
    }
  }

  /**
   * Handles edits made to the 'Category' column in the 'Transactions' sheet.
   * It updates the 'Sub-Category' dropdown options based on the selected 'Type' and 'Category',
   * and manages UI highlights.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The 'Transactions' sheet object.
   * @param {number} row - The 1-based row number of the cell that was edited.
   * @param {Object<string, Array<string>>} typeCategoryToSubCategories - The mapping of (Type+Category)
   *   composite keys to their available sub-categories, obtained from the dropdown cache.
   * @private
   */
  function handleCategoryEdit(sheet, row, typeCategoryToSubCategories) {
    const typeCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.TYPE);
    const categoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.CATEGORY);
    const subCategoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.SUB_CATEGORY);

    const type = typeCell.getValue();
    const category = categoryCell.getValue();
    const oldSubCategoryValue = subCategoryCell.getValue();

    if (type && category && category !== DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT) {
      const key = `${type}${DROPDOWN_CONFIG.KEY_SEPARATOR}${category}`;
      const subCategories = typeCategoryToSubCategories[key] || [];
      setDropdownValidation(subCategoryCell, subCategories);

      if (oldSubCategoryValue && (subCategories.length === 0 || !subCategories.includes(oldSubCategoryValue))) {
        subCategoryCell.clearContent();
        highlightPending(subCategoryCell);
      } else if (subCategories.length > 0 && !oldSubCategoryValue) {
        highlightPending(subCategoryCell); // Highlight if new options available and cell is empty
      }

    } else {
      // If type or category is placeholder or empty, clear sub-category
      subCategoryCell.clearContent().clearDataValidations();
    }

    clearPlaceholderIfNeeded(categoryCell);
    if (subCategoryCell.getValue() !== DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT && subCategoryCell.getValue() !== "") {
        clearHighlight(subCategoryCell);
    }
  }

  /**
   * Handles edits made to the 'Sub-Category' column in the 'Transactions' sheet.
   * Primarily clears any placeholder text from the cell and removes pending highlights.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The 'Transactions' sheet object.
   * @param {number} row - The 1-based row number of the cell that was edited.
   * @private
   */
  function handleSubCategoryEdit(sheet, row) {
    const subCategoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.SUB_CATEGORY);
    clearPlaceholderIfNeeded(subCategoryCell);
    clearHighlight(subCategoryCell); // Clear highlight once a sub-category is selected or cleared
  }


  // Public API
  return {
    /**
     * Handles edit events on the spreadsheet, specifically for the 'Transactions' sheet.
     * This function is intended to be called by an `onEdit` trigger.
     * It updates dependent dropdowns for 'Category' and 'Sub-Category' based on the edited cell
     * in the 'Transactions' sheet. It handles single cell edits and multi-cell pastes.
     * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The edit event object provided by Google Apps Script.
     * @memberof FinancialPlanner.DropdownService
     * @return {void}
     *
     * @example
     * // Typically called from a global onEdit trigger:
     * // function onEdit(e) {
     * //   if (e.range.getSheet().getName() === FinancialPlanner.Config.get().SHEETS.TRANSACTIONS) {
     * //     FinancialPlanner.DropdownService.handleEdit(e);
     * //   }
     * //   // ... other onEdit logic
     * // }
     */
    handleEdit: function(e) {
      try {
        const sheet = e.range.getSheet();
        const transactionsSheetName = (config && config.get && config.get().SHEETS && config.get().SHEETS.TRANSACTIONS)
                                      ? config.get().SHEETS.TRANSACTIONS
                                      : DROPDOWN_CONFIG.SHEETS.TRANSACTIONS;

        if (sheet.getName() !== transactionsSheetName) return;
        // Skip header row if it's frozen and the edit is in the first row
        if (e.range.getRow() === 1 && sheet.getFrozenRows() >= 1) return;

        // Check if sheet has expected structure for dropdowns (Type/Category/SubCategory columns)
        // If headers don't match expected structure, skip dropdown processing
        if (sheet.getLastRow() > 0) {
          const headerRow = sheet.getRange(1, 1, 1, Math.min(6, sheet.getLastColumn())).getValues()[0];
          const hasExpectedStructure = headerRow[2] === 'Type' || headerRow[2] === 'Category'; // Column C or D
          if (!hasExpectedStructure) {
            Logger.log("DropdownService: Sheet structure doesn't match expected format. Skipping dropdown processing.");
            return;
          }
        }

        if (!dropdownCache) {
          Logger.log("DropdownService: Initializing dropdown cache in handleEdit.");
          dropdownCache = buildDropdownCache(e.source); // e.source is the Spreadsheet
        }
        const { typeToCategories, typeCategoryToSubCategories } = dropdownCache;

        const startRow = e.range.getRow();
        const startCol = e.range.getColumn();
        const numRows = e.range.getNumRows();
        const numCols = e.range.getNumColumns();

        // Iterate over the edited range (can be multiple cells if pasted)
        for (let r = 0; r < numRows; r++) {
          const currentRow = startRow + r;
          // Skip header row again if iterating over a multi-row paste that includes the header
          if (currentRow === 1 && sheet.getFrozenRows() >= 1) continue;

          for (let c = 0; c < numCols; c++) {
            const currentCol = startCol + c;

            if (currentCol === DROPDOWN_CONFIG.COLUMNS.TYPE) {
              handleTypeEdit(sheet, currentRow, typeToCategories);
            } else if (currentCol === DROPDOWN_CONFIG.COLUMNS.CATEGORY) {
              handleCategoryEdit(sheet, currentRow, typeCategoryToSubCategories);
            } else if (currentCol === DROPDOWN_CONFIG.COLUMNS.SUB_CATEGORY) {
              handleSubCategoryEdit(sheet, currentRow);
            }
          }
        }
      } catch (error) {
        Logger.log('DropdownService: Error in handleEdit: ' + error.toString());
        errorService.handle(errorService.create('Error in dropdown edit handler', { originalError: error.toString(), eventDetails: e ? JSON.stringify(e) : 'N/A' }), "An error occurred while updating dropdowns.");
      }
    },

    /**
     * Forces a refresh of the dropdown cache. It clears the relevant key from Google Apps Script's
     * `CacheService` and then rebuilds the in-memory `dropdownCache` by reading from the
     * 'Dropdowns' sheet. Provides UI feedback (loading spinner, success/error notifications).
     * @memberof FinancialPlanner.DropdownService
     * @return {void}
     *
     * @example
     * FinancialPlanner.DropdownService.refreshCache();
     */
    refreshCache: function() {
      try {
        uiService.showLoadingSpinner("Refreshing dropdown options...");
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const scriptCache = CacheService.getScriptCache(); // Native Google Apps Script CacheService
        scriptCache.remove(DROPDOWN_CONFIG.CACHE_KEY);
        Logger.log("DropdownService: ScriptCache cleared for dropdowns.");
        dropdownCache = buildDropdownCache(spreadsheet); // Rebuild in-memory cache
        uiService.hideLoadingSpinner();
        uiService.showSuccessNotification('Dropdown options have been refreshed.');
      } catch (error) {
        uiService.hideLoadingSpinner();
        Logger.log('DropdownService: Error refreshing dropdown cache: ' + error.toString());
        errorService.handle(errorService.create('Error refreshing dropdown cache', { originalError: error.toString() }), "Failed to refresh dropdown options.");
      }
    },

    /**
     * Initializes the in-memory `dropdownCache` by calling `buildDropdownCache` if it's not already loaded.
     * This can be useful for pre-loading the cache (e.g., during an `onOpen` event)
     * in scenarios where `handleEdit` might not be the first function to require it.
     * @memberof FinancialPlanner.DropdownService
     * @return {void}
     *
     * @example
     * FinancialPlanner.DropdownService.initializeCache();
     */
    initializeCache: function() {
        if (!dropdownCache) {
            Logger.log("DropdownService: Initializing dropdown cache proactively.");
            dropdownCache = buildDropdownCache(SpreadsheetApp.getActiveSpreadsheet());
        }
    }
  };
})(FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.ErrorService, FinancialPlanner.Config);

// Backward compatibility layer for existing global functions

/**
 * Global `onEdit` trigger function for Google Apps Script.
 * This version primarily delegates to `FinancialPlanner.DropdownService.handleEdit(e)`
 * if the edit occurs on the configured 'Transactions' sheet.
 * For a more robust solution with multiple modules needing `onEdit` handling, a central event dispatcher
 * (e.g., `FinancialPlanner.Controllers.onEdit`) should be used to route the event.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event object provided by Google Apps Script.
 * @global
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  let transactionsSheetName = 'Transactions'; // Default
  try {
    if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Config && FinancialPlanner.Config.get) {
      transactionsSheetName = FinancialPlanner.Config.get().SHEETS.TRANSACTIONS || 'Transactions';
    }
  } catch (configError) {
    Logger.log("Global onEdit: Error accessing FinancialPlanner.Config: " + configError.toString());
  }

  if (sheet.getName() === transactionsSheetName) {
    if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.DropdownService && FinancialPlanner.DropdownService.handleEdit) {
      FinancialPlanner.DropdownService.handleEdit(e);
    } else {
      Logger.log("Global onEdit: FinancialPlanner.DropdownService.handleEdit not available for delegation.");
    }
  }
  // Consider calling a central dispatcher if other modules need onEdit:
  // if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Controllers && FinancialPlanner.Controllers.onEdit) {
  //   FinancialPlanner.Controllers.onEdit(e); // This would then call DropdownService.handleEdit and others
  // }
}


/**
 * Global function to manually refresh the dropdown cache.
 * This function delegates to `FinancialPlanner.DropdownService.refreshCache()`.
 * It can be called from the Google Sheets UI (e.g., via a custom menu).
 * @global
 */
function refreshCache() {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.DropdownService && FinancialPlanner.DropdownService.refreshCache) {
    FinancialPlanner.DropdownService.refreshCache();
  } else {
    Logger.log("Global refreshCache: FinancialPlanner.DropdownService.refreshCache not available.");
    try {
      SpreadsheetApp.getUi().alert("Could not refresh dropdown options: The DropdownService is not loaded correctly.");
    } catch (uiError) {
      Logger.log("Global refreshCache: Failed to show UI alert: " + uiError.toString());
    }
  }
}
