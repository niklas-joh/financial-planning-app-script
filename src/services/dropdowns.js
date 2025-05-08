/**
 * Financial Planning Tools - Dropdown Service
 *
 * This file provides dynamic dependent dropdown functionality for the Transactions sheet.
 * It follows the namespace pattern and uses dependency injection for better maintainability.
 */

// Create the DropdownService module within the FinancialPlanner namespace
FinancialPlanner.DropdownService = (function(utils, uiService, errorService, config) {
  // Private constants
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
  let dropdownCache = null;

  /**
   * Converts Set values in an object to Arrays.
   * @param {Object} obj - Object with Set values.
   * @return {Object} Object with Array values.
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
   * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The active spreadsheet.
   * @return {Object} Cache object with dropdown mappings.
   * @private
   */
  function buildDropdownCache(spreadsheet) {
    try {
      // Try to get from script cache first
      const cache = CacheService.getScriptCache(); // Google Apps Script Cache Service
      let cachedData = cache.get(DROPDOWN_CONFIG.CACHE_KEY);

      if (cachedData) {
        Logger.log('Using cached dropdown data');
        const parsed = JSON.parse(cachedData);
        // Ensure the structure matches what's expected
        return {
          typeToCategories: parsed.typeToCategories || {},
          typeCategoryToSubCategories: parsed.typeCategoryToSubCategories || {},
        };
      }

      Logger.log('Building dropdown cache from sheet');
      const dropdownsSheetName = DROPDOWN_CONFIG.SHEETS.DROPDOWNS; // Use local config
      const dropdownsSheet = spreadsheet.getSheetByName(dropdownsSheetName);

      if (!dropdownsSheet) {
        throw new Error(`Sheet "${dropdownsSheetName}" not found`);
      }

      const dropdownsData = dropdownsSheet.getDataRange().getValues();
      const startRow = (dropdownsData.length > 0 &&
                       typeof dropdownsData[0][0] === 'string' &&
                       dropdownsData[0][0].toLowerCase() === 'type') ? 1 : 0;

      const typeToCategories = {};
      const typeCategoryToSubCategories = {};

      for (let i = startRow; i < dropdownsData.length; i++) {
        const [type, category, subCategory] = dropdownsData[i];
        if (!type || !category) continue;

        if (!typeToCategories[type]) {
          typeToCategories[type] = new Set();
        }
        typeToCategories[type].add(category);

        if (subCategory) {
          const key = `${type}${DROPDOWN_CONFIG.KEY_SEPARATOR}${category}`;
          if (!typeCategoryToSubCategories[key]) {
            typeCategoryToSubCategories[key] = new Set();
          }
          typeCategoryToSubCategories[key].add(subCategory);
        }
      }

      const cacheToStore = {
        typeToCategories: mapSetsToArrays(typeToCategories),
        typeCategoryToSubCategories: mapSetsToArrays(typeCategoryToSubCategories),
      };

      try {
        cache.put(DROPDOWN_CONFIG.CACHE_KEY, JSON.stringify(cacheToStore), DROPDOWN_CONFIG.CACHE_EXPIRY_SECONDS);
      } catch (e) {
        Logger.log('Failed to cache dropdown data: ' + e.toString());
      }

      return cacheToStore;
    } catch (error) {
      Logger.log('Error building dropdown cache: ' + error.toString());
      errorService.log(errorService.create('Error building dropdown cache', { originalError: error.toString() }));
      return {
        typeToCategories: {},
        typeCategoryToSubCategories: {}
      };
    }
  }

  /**
   * Sets up a dropdown validation in a cell.
   * @param {SpreadsheetApp.Range} range - The cell range to set the dropdown in.
   * @param {Array} options - The dropdown options.
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
   * Clears the placeholder text if it's present in a cell.
   * @param {SpreadsheetApp.Range} range - The cell range to check.
   * @private
   */
  function clearPlaceholderIfNeeded(range) {
    if (range.getValue() === DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT) {
      range.clearContent();
    }
  }

  /**
   * Highlights a cell to indicate pending selection.
   * @param {SpreadsheetApp.Range} range - The cell range to highlight.
   * @private
   */
  function highlightPending(range) {
    range.setBackground(DROPDOWN_CONFIG.UI.PENDING_BACKGROUND);
  }

  /**
   * Clears highlighting from a cell.
   * @param {SpreadsheetApp.Range} range - The cell range to clear highlighting from.
   * @private
   */
  function clearHighlight(range) {
    range.setBackground(null);
  }

  /**
   * Handles edits to the Type column.
   * @param {SpreadsheetApp.Sheet} sheet - The active sheet.
   * @param {number} row - The edited row.
   * @param {Object} typeToCategories - Mapping of types to categories.
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
   * Handles edits to the Category column.
   * @param {SpreadsheetApp.Sheet} sheet - The active sheet.
   * @param {number} row - The edited row.
   * @param {Object} typeCategoryToSubCategories - Mapping of type+category to subcategories.
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
   * Handles edits to the Sub-Category column.
   * @param {SpreadsheetApp.Sheet} sheet - The active sheet.
   * @param {number} row - The edited row.
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
     * Handles edit events in the spreadsheet, specifically for the Transactions sheet.
     * @param {Object} e - The edit event object.
     */
    handleEdit: function(e) {
      try {
        const sheet = e.range.getSheet();
        // Use global config for sheet names if available, otherwise fallback to local
        const transactionsSheetName = (config && config.get && config.get().SHEETS && config.get().SHEETS.TRANSACTIONS)
                                      ? config.get().SHEETS.TRANSACTIONS
                                      : DROPDOWN_CONFIG.SHEETS.TRANSACTIONS;

        if (sheet.getName() !== transactionsSheetName) return;
        if (e.range.getRow() === 1 && sheet.getFrozenRows() >= 1) return; // Skip header row if frozen

        if (!dropdownCache) {
          dropdownCache = buildDropdownCache(e.source);
        }
        const { typeToCategories, typeCategoryToSubCategories } = dropdownCache;

        const startRow = e.range.getRow();
        const startCol = e.range.getColumn();
        const numRows = e.range.getNumRows();
        const numCols = e.range.getNumColumns();

        for (let r = 0; r < numRows; r++) {
          const currentRow = startRow + r;
          if (currentRow === 1 && sheet.getFrozenRows() >= 1) continue; // Skip header again if iterating

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
        Logger.log('Error in DropdownService.handleEdit: ' + error.toString());
        errorService.handle(errorService.create('Error in dropdown edit handler', { originalError: error.toString() }), "An error occurred while updating dropdowns.");
      }
    },

    /**
     * Forces a refresh of the dropdown cache.
     */
    refreshCache: function() {
      try {
        uiService.showLoadingSpinner("Refreshing dropdown cache...");
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const scriptCache = CacheService.getScriptCache();
        scriptCache.remove(DROPDOWN_CONFIG.CACHE_KEY);
        dropdownCache = buildDropdownCache(spreadsheet); // Rebuild in-memory cache as well
        uiService.hideLoadingSpinner();
        uiService.showSuccessNotification('Dropdown cache has been refreshed.');
      } catch (error) {
        uiService.hideLoadingSpinner();
        Logger.log('Error refreshing dropdown cache: ' + error.toString());
        errorService.handle(errorService.create('Error refreshing dropdown cache', { originalError: error.toString() }), "Failed to refresh dropdown cache.");
      }
    },

    /**
     * Initializes the dropdown cache if it's not already loaded.
     * Useful for scenarios where onEdit might not be the first trigger.
     */
    initializeCache: function() {
        if (!dropdownCache) {
            Logger.log("Initializing dropdown cache proactively.");
            dropdownCache = buildDropdownCache(SpreadsheetApp.getActiveSpreadsheet());
        }
    }
  };
})(FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.ErrorService, FinancialPlanner.Config);

// Backward compatibility layer for existing global functions

/**
 * Global onEdit trigger.
 * It now delegates to the appropriate service if the edit is on the Transactions sheet.
 * Other onEdit functionalities for different sheets should be handled by FinancialPlanner.EventHandlers.onEdit
 * @param {Object} e The event object
 */
function onEdit(e) {
  // This global onEdit should ideally be managed by a central event dispatcher.
  // For now, we check if the edit is on the Transactions sheet and delegate.
  const sheet = e.range.getSheet();
  // Use FinancialPlanner.Config if available for sheet name
  let transactionsSheetName = 'Transactions'; // Default
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.Config && FinancialPlanner.Config.get) {
    transactionsSheetName = FinancialPlanner.Config.get().SHEETS.TRANSACTIONS || 'Transactions';
  }

  if (sheet.getName() === transactionsSheetName) {
    if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.DropdownService) {
      FinancialPlanner.DropdownService.handleEdit(e);
    } else {
      Logger.log("FinancialPlanner.DropdownService not available for onEdit delegation.");
    }
  }
  // IMPORTANT: If other modules also need onEdit, a central dispatcher in FinancialPlanner.EventHandlers
  // should be responsible for calling the respective module handlers.
  // For example: FinancialPlanner.EventHandlers.onEdit(e); which then calls DropdownService.handleEdit(e) etc.
}


/**
 * Global function to manually refresh the dropdown cache.
 * Delegates to the DropdownService.
 */
function refreshCache() {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.DropdownService) {
    return FinancialPlanner.DropdownService.refreshCache();
  }
  Logger.log("FinancialPlanner.DropdownService not available for refreshCache delegation.");
  SpreadsheetApp.getUi().alert("Could not refresh cache: DropdownService not loaded.");
}
