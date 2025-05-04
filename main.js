/**
 * Google Apps Script to create dynamic dependent dropdowns (Type -> Category -> Sub-Category)
 * Based on sheets:
 *   - 'Transactions' (where dropdowns are applied)
 *   - 'Dropdowns' (mapping table with Type, Category, Sub-Category)
 * Enhanced with:
 *   - Caching for improved performance
 *   - Adding "Please select" as default option
 *   - Allowing users to enter values outside the dropdown list
 *   - Auto-clear 'Please select' if accidentally saved
 *   - Proper handling of Ctrl+D and multi-cell edits
 *   - Grey out columns while awaiting user action
 *   - Load mapping into memory for faster performance
 */

// ===== CONSTANTS =====
const CONFIG = {
  CACHE_EXPIRY_SECONDS: 300, // Cache expires in 5 minutes
  CACHE_KEY: 'dropdownsData',
  SHEETS: {
    TRANSACTIONS: 'Transactions',
    DROPDOWNS: 'Dropdowns'
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

// ===== CACHE MANAGEMENT =====
let GLOBAL_DROPDOWN_CACHE = null;

/**
 * Converts Set values in an object to Arrays
 * @param {Object} obj - Object with Set values
 * @return {Object} Object with Array values
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
 * Builds or retrieves the dropdown cache
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The active spreadsheet
 * @return {Object} Cache object with dropdown mappings
 */
function buildDropdownCache(spreadsheet) {
  try {
    // Try to get from script cache first
    const cache = CacheService.getScriptCache();
    let cachedData = cache.get(CONFIG.CACHE_KEY);

    if (cachedData) {
      Logger.log('Using cached dropdown data');
      const parsed = JSON.parse(cachedData);
      return {
        typeToCategories: parsed.typeToCategories,
        typeCategoryToSubCategories: parsed.typeCategoryToSubCategories,
      };
    }

    // If not in cache, build from sheet
    Logger.log('Building dropdown cache from sheet');
    const dropdownsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DROPDOWNS);
    
    if (!dropdownsSheet) {
      throw new Error(`Sheet "${CONFIG.SHEETS.DROPDOWNS}" not found`);
    }
    
    const dropdownsData = dropdownsSheet.getDataRange().getValues();
    
    // Skip header row if it exists (check if first row contains headers)
    const startRow = (dropdownsData.length > 0 && 
                     typeof dropdownsData[0][0] === 'string' && 
                     dropdownsData[0][0].toLowerCase() === 'type') ? 1 : 0;
    
    const typeToCategories = {};
    const typeCategoryToSubCategories = {};

    // Process each row in the dropdowns data
    for (let i = startRow; i < dropdownsData.length; i++) {
      const [type, category, subCategory] = dropdownsData[i];
      
      // Skip empty rows or rows with missing data
      if (!type || !category) continue;
      
      // Add to type -> categories mapping
      if (!typeToCategories[type]) {
        typeToCategories[type] = new Set();
      }
      typeToCategories[type].add(category);

      // Add to type+category -> subcategories mapping
      if (subCategory) {
        const key = `${type}${CONFIG.KEY_SEPARATOR}${category}`;
        if (!typeCategoryToSubCategories[key]) {
          typeCategoryToSubCategories[key] = new Set();
        }
        typeCategoryToSubCategories[key].add(subCategory);
      }
    }

    // Convert Sets to Arrays for caching
    const cacheData = {
      typeToCategories: mapSetsToArrays(typeToCategories),
      typeCategoryToSubCategories: mapSetsToArrays(typeCategoryToSubCategories),
    };

    // Store in cache
    try {
      cache.put(CONFIG.CACHE_KEY, JSON.stringify(cacheData), CONFIG.CACHE_EXPIRY_SECONDS);
    } catch (e) {
      Logger.log('Failed to cache dropdown data: ' + e.toString());
      // Continue even if caching fails
    }

    return cacheData;
  } catch (error) {
    Logger.log('Error building dropdown cache: ' + error.toString());
    // Return empty cache structure on error
    return {
      typeToCategories: {},
      typeCategoryToSubCategories: {}
    };
  }
}

/**
 * Forces a refresh of the dropdown cache
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The active spreadsheet
 */
function refreshDropdownCache(spreadsheet) {
  const cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE_KEY);
  GLOBAL_DROPDOWN_CACHE = buildDropdownCache(spreadsheet);
}

// ===== UI HELPERS =====

/**
 * Sets up a dropdown validation in a cell
 * @param {Range} range - The cell range to set the dropdown in
 * @param {Array} options - The dropdown options
 */
function setDropdown(range, options) {
  range.clearDataValidations();
  
  if (!options || options.length === 0) {
    return;
  }
  
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.UI.PLACEHOLDER_TEXT, ...options], true)
    .setAllowInvalid(true)
    .build();
    
  range.setDataValidation(rule);
}

/**
 * Clears the placeholder text if it's present in a cell
 * @param {Range} range - The cell range to check
 */
function clearPlaceholderIfNeeded(range) {
  if (range.getValue() === CONFIG.UI.PLACEHOLDER_TEXT) {
    range.clearContent();
  }
}

/**
 * Highlights a cell to indicate pending selection
 * @param {Range} range - The cell range to highlight
 */
function highlightPending(range) {
  range.setBackground(CONFIG.UI.PENDING_BACKGROUND);
}

/**
 * Clears highlighting from a cell
 * @param {Range} range - The cell range to clear highlighting from
 */
function clearHighlight(range) {
  range.setBackground(null);
}

// ===== MAIN EDIT HANDLER =====

/**
 * Handles edit events in the spreadsheet
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  try {
    // Check if edit is in the Transactions sheet
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEETS.TRANSACTIONS) return;
    
    // Skip header row if present (assuming row 1 is header)
    if (e.range.getRow() === 1) return;
    
    // Initialize cache if needed
    if (!GLOBAL_DROPDOWN_CACHE) {
      GLOBAL_DROPDOWN_CACHE = buildDropdownCache(e.source);
    }
    
    const { typeToCategories, typeCategoryToSubCategories } = GLOBAL_DROPDOWN_CACHE;

    // Get the range dimensions to handle multi-cell edits (like Ctrl+D)
    const startRow = e.range.getRow();
    const startCol = e.range.getColumn();
    const numRows = e.range.getNumRows();
    const numCols = e.range.getNumColumns();
    
    // Process each cell in the edited range
    for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
      const currentRow = startRow + rowOffset;
      
      // Skip header row
      if (currentRow === 1) continue;
      
      for (let colOffset = 0; colOffset < numCols; colOffset++) {
        const currentCol = startCol + colOffset;
        
        // Handle Type column edits
        if (currentCol === CONFIG.COLUMNS.TYPE) {
          handleTypeEdit(sheet, currentRow, typeToCategories);
        }
        
        // Handle Category column edits
        else if (currentCol === CONFIG.COLUMNS.CATEGORY) {
          handleCategoryEdit(sheet, currentRow, typeCategoryToSubCategories);
        }
        
        // Handle Sub-Category column edits
        else if (currentCol === CONFIG.COLUMNS.SUB_CATEGORY) {
          handleSubCategoryEdit(sheet, currentRow);
        }
      }
    }
  } catch (error) {
    Logger.log('Error in onEdit: ' + error.toString());
    // We don't want to throw errors in onEdit as it would disrupt the user
  }
}

/**
 * Handles edits to the Type column
 * @param {Sheet} sheet - The active sheet
 * @param {Number} row - The edited row
 * @param {Object} typeToCategories - Mapping of types to categories
 */
function handleTypeEdit(sheet, row, typeToCategories) {
  const typeCell = sheet.getRange(row, CONFIG.COLUMNS.TYPE);
  const categoryCell = sheet.getRange(row, CONFIG.COLUMNS.CATEGORY);
  const subCategoryCell = sheet.getRange(row, CONFIG.COLUMNS.SUB_CATEGORY);
  
  // Get the new type value
  const type = typeCell.getValue();
  
  // Only clear dependent fields if this is a direct edit, not a Ctrl+D operation
  // We can detect this by checking if the value is new or changed
  const oldCategoryValue = categoryCell.getValue();
  const oldSubCategoryValue = subCategoryCell.getValue();
  
  // Check if we need to update the category dropdown
  const categories = typeToCategories[type] || [];
  
  // Set up category dropdown based on selected type
  setDropdown(categoryCell, categories);
  
  // Only clear category content if the type has changed and would make the current category invalid
  if (oldCategoryValue && categories.length > 0 && !categories.includes(oldCategoryValue)) {
    categoryCell.clearContent();
    subCategoryCell.clearContent().clearDataValidations();
    highlightPending(categoryCell);
  }
  
  // Clear placeholder if needed
  clearPlaceholderIfNeeded(typeCell);
  clearHighlight(categoryCell);
}

/**
 * Handles edits to the Category column
 * @param {Sheet} sheet - The active sheet
 * @param {Number} row - The edited row
 * @param {Object} typeCategoryToSubCategories - Mapping of type+category to subcategories
 */
function handleCategoryEdit(sheet, row, typeCategoryToSubCategories) {
  const typeCell = sheet.getRange(row, CONFIG.COLUMNS.TYPE);
  const categoryCell = sheet.getRange(row, CONFIG.COLUMNS.CATEGORY);
  const subCategoryCell = sheet.getRange(row, CONFIG.COLUMNS.SUB_CATEGORY);
  
  // Get the current type and category values
  const type = typeCell.getValue();
  const category = categoryCell.getValue();
  const oldSubCategoryValue = subCategoryCell.getValue();
  
  // Only proceed if we have both type and category
  if (type && category) {
    const key = `${type}${CONFIG.KEY_SEPARATOR}${category}`;
    const subCategories = typeCategoryToSubCategories[key] || [];
    
    // Set up sub-category dropdown
    setDropdown(subCategoryCell, subCategories);
    
    // Only clear sub-category content if the category has changed and would make the current sub-category invalid
    if (oldSubCategoryValue && subCategories.length > 0 && !subCategories.includes(oldSubCategoryValue)) {
      subCategoryCell.clearContent();
      highlightPending(subCategoryCell);
    }
  }
  
  // Clear placeholder if needed
  clearPlaceholderIfNeeded(categoryCell);
  clearHighlight(subCategoryCell);
}

/**
 * Handles edits to the Sub-Category column
 * @param {Sheet} sheet - The active sheet
 * @param {Number} row - The edited row
 */
function handleSubCategoryEdit(sheet, row) {
  const subCategoryCell = sheet.getRange(row, CONFIG.COLUMNS.SUB_CATEGORY);
  clearPlaceholderIfNeeded(subCategoryCell);
}

/**
 * Menu function to manually refresh the dropdown cache
 */
function refreshCache() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  refreshDropdownCache(spreadsheet);
  SpreadsheetApp.getUi().alert('Dropdown cache has been refreshed.');
}

/**
 * Creates a custom menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Dropdown Tools')
    .addItem('Refresh Dropdown Cache', 'refreshCache')
    .addToUi();
}
