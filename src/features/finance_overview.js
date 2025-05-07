/**
 * Financial Planning Tools - Financial Overview Generator
 * 
 * This module creates a comprehensive financial overview sheet based on transaction data.
 * It generates a complete overview with dynamic categories and optional sub-category display
 * based on user preference.
 * 
 * Version: 2.0.0
 * Last Updated: 2025-05-06
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Configuration specific to the financial overview functionality
 * Namespace-specific prefix prevents conflicts with other modules
 */
const FINANCE_OVERVIEW_CONFIG = {
  SHEETS: {
    OVERVIEW: "Overview",
    TRANSACTIONS: "Transactions",
    DROPDOWNS: "Dropdowns",
    ERROR_LOG: "Error Log",
    ANALYSIS: "Analysis"  // New sheet for dedicated analysis
  },
  TRANSACTION_TYPES: {
    INCOME: "Income",
    ESSENTIALS: "Essentials",
    WANTS: "Wants/Pleasure",
    EXTRA: "Extra", 
    SAVINGS: "Savings"
  },
  // Transaction types in preferred display order
  TYPE_ORDER: ["Income", "Essentials", "Wants/Pleasure", "Extra", "Savings"],
  // Types considered as expenses (for shared expense calculations)
  EXPENSE_TYPES: ["Essentials", "Wants/Pleasure", "Extra"],
  // Default target rates for expense categories
  TARGET_RATES: {
    ESSENTIALS: 0.5,    // 50% for essentials
    WANTS: 0.3,         // 30% for wants/pleasure
    EXTRA: 0.2,         // 20% for extras
    DEFAULT: 0.2        // 20% default
  },
  // Header structure for the overview sheet
  HEADERS: [
    "Type", "Category", "Sub-Category", "Shared?", 
    "Jan-24", "Feb-24", "Mar-24", "Apr-24", 
    "May-24", "Jun-24", "Jul-24", "Aug-24", 
    "Sep-24", "Oct-24", "Nov-24", "Dec-24", "Total", "Average"
  ],
  // UI element positions and names
  UI: {
    SUBCATEGORY_TOGGLE: {
      LABEL_CELL: "S1",
      CHECKBOX_CELL: "T1",
      LABEL_TEXT: "Show Sub-Categories",
      NOTE_TEXT: "Toggle to show or hide sub-categories in the overview sheet"
    },
    COLUMN_WIDTHS: {
      TYPE: 150,
      CATEGORY: 150,
      SUBCATEGORY: 150,
      SHARED: 80,
      MONTH: 90,
      AVERAGE: 100,
      EXPENSE_CATEGORY: 150,
      AMOUNT: 100,
      RATE: 80
    }
  },
  COLORS: {
    // Type header colors
    TYPE_HEADERS: {
      INCOME: {
        BG: "#2E7D32",      // Green for Income
        FONT: "#FFFFFF"     // White text
      },
      ESSENTIALS: {
        BG: "#1976D2",      // Blue for Essentials
        FONT: "#FFFFFF"     // White text
      },
      WANTS_PLEASURE: {
        BG: "#FFA000",      // Amber for Wants/Pleasure
        FONT: "#FFFFFF"     // White text
      },
      EXTRA: {
        BG: "#7B1FA2",      // Purple for Extra
        FONT: "#FFFFFF"     // White text
      },
      SAVINGS: {
        BG: "#1565C0",      // Blue for Savings
        FONT: "#FFFFFF"     // White text
      },
      DEFAULT: {
        BG: "#424242",      // Dark gray for other types
        FONT: "#FFFFFF"     // White text
      }
    },
    // UI element colors
    UI: {
      HEADER_BG: "#C62828",       // Deep red for headers
      HEADER_FONT: "#FFFFFF",     // White text for headers
      METRICS_BG: "#FFEBEE",      // Very light red for metrics section
      BORDER: "#FF8F00",          // Amber for borders
      INCOME_FONT: "#388E3C",     // Green for income values
      EXPENSE_FONT: "#D32F2F",    // Red for expense values
      SAVINGS_FONT: "#1565C0",    // Blue for savings values
      NEUTRAL_FONT: "#424242",    // Dark gray for neutral values
      NET_BG: "#424242",          // Dark gray for net calculations
      NET_FONT: "#FFFFFF"         // White text for net calculations
    },
    // Chart colors
    CHART: {
      SERIES: [
        "#C62828", // Red (for Essentials)
        "#FF8F00", // Amber (for Wants/Pleasure)
        "#1565C0", // Blue (for Extra)
        "#2E7D32", // Green
        "#6A1B9A", // Purple
        "#E64A19", // Deep Orange
        "#00695C", // Teal
        "#5D4037"  // Brown
      ],
      TITLE: "#424242",
      TEXT: "#424242"
    }
  },
  // Cache configuration
  CACHE: {
    ENABLED: true,
    KEYS: {
      CATEGORY_COMBINATIONS: "finance_overview_categories",
      GROUPED_COMBINATIONS: "finance_overview_grouped"
    },
    EXPIRY_SECONDS: 21600 // 6 hours
  },
  // Locale settings
  LOCALE: {
    CURRENCY_SYMBOL: "€",
    CURRENCY_LOCALE: "2", // Euro
    DATE_FORMAT: "yyyy-MM-dd"
  },
  // Performance settings
  PERFORMANCE: {
    BATCH_SIZE: 50, // Number of rows to process in one batch
    USE_BATCH_OPERATIONS: true
  }
};

// ============================================================================
// CUSTOM ERROR HANDLING
// ============================================================================

/**
 * Custom error class for Financial Overview operations
 * @class
 * @extends Error
 */
class FinanceOverviewError extends Error {
  /**
   * Creates a new FinanceOverviewError
   * @param {String} message - Error message
   * @param {Object} details - Additional error details
   */
  constructor(message, details = {}) {
    super(message);
    this.name = 'FinanceOverviewError';
    this.details = details;
    this.timestamp = new Date();
  }
  
  /**
   * Logs the error to a dedicated error log sheet
   */
  logToSheet() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const errorSheet = getOrCreateSheet(ss, FINANCE_OVERVIEW_CONFIG.SHEETS.ERROR_LOG);
      
      // Create headers if this is a new sheet
      if (errorSheet.getLastRow() === 0) {
        errorSheet.appendRow(["Timestamp", "Error Type", "Message", "Details"]);
        errorSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
      }
      
      // Append error information
      errorSheet.appendRow([
        this.timestamp, 
        this.name, 
        this.message, 
        JSON.stringify(this.details)
      ]);
      
      // Format the timestamp
      const lastRow = errorSheet.getLastRow();
      errorSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
      
      // Set colors based on error severity
      const bgColor = this.details.severity === "high" ? "#F9BDBD" : 
                      this.details.severity === "medium" ? "#FFE0B2" : "#E1F5FE";
      errorSheet.getRange(lastRow, 1, 1, 4).setBackground(bgColor);
    } catch (error) {
      // If we can't log to sheet, at least log to console
      console.error("Failed to log error to sheet:", error);
      console.error("Original error:", this.message, this.details);
    }
  }
}

// ============================================================================
// CACHING UTILITY
// ============================================================================

/**
 * Utility for caching expensive operations
 */
const CacheUtil = {
  /**
   * Gets a value from cache or computes it if not available
   * @param {String} key - Cache key
   * @param {Function} computeFunction - Function to compute value if not in cache
   * @param {Number} expirySeconds - Cache expiry in seconds
   * @return {any} The cached or computed value
   */
  getCachedOrCompute(key, computeFunction, expirySeconds = FINANCE_OVERVIEW_CONFIG.CACHE.EXPIRY_SECONDS) {
    if (!FINANCE_OVERVIEW_CONFIG.CACHE.ENABLED) {
      return computeFunction();
    }
    
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(key);
      
      if (cached != null) {
        return JSON.parse(cached);
      }
      
      const result = computeFunction();
      
      try {
        cache.put(key, JSON.stringify(result), expirySeconds);
      } catch (cacheError) {
        console.warn(`Failed to cache result for key ${key}:`, cacheError);
      }
      
      return result;
    } catch (error) {
      console.warn(`Cache operation failed for key ${key}:`, error);
      // Fall back to direct computation
      return computeFunction();
    }
  },
  
  /**
   * Invalidates a specific cache entry
   * @param {String} key - Cache key to invalidate
   */
  invalidate(key) {
    if (!FINANCE_OVERVIEW_CONFIG.CACHE.ENABLED) return;
    
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(key);
    } catch (error) {
      console.warn(`Failed to invalidate cache for key ${key}:`, error);
    }
  },
  
  /**
   * Invalidates all finance overview cache entries
   */
  invalidateAll() {
    if (!FINANCE_OVERVIEW_CONFIG.CACHE.ENABLED) return;
    
    try {
      const cache = CacheService.getScriptCache();
      const keys = Object.values(FINANCE_OVERVIEW_CONFIG.CACHE.KEYS);
      cache.removeAll(keys);
    } catch (error) {
      console.warn("Failed to invalidate all cache entries:", error);
    }
  }
};

// ============================================================================
// USER INTERFACE UTILITIES
// ============================================================================

/**
 * Utilities for enhanced user interface feedback
 */
const UIUtil = {
  /**
   * Shows a loading spinner with a message
   * @param {String} message - Message to display
   */
  showLoadingSpinner(message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Working...");
  },
  
  /**
   * Hides the loading spinner
   */
  hideLoadingSpinner() {
    // Google Apps Script doesn't have a direct way to dismiss toasts
    // So we just show a blank toast that will disappear quickly
    SpreadsheetApp.getActiveSpreadsheet().toast("", "", 1);
  },
  
  /**
   * Shows a success notification
   * @param {String} message - Success message
   * @param {Number} duration - Duration in seconds to show the message
   */
  showSuccessNotification(message, duration = 5) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Success", duration);
  },
  
  /**
   * Shows an error notification
   * @param {String} title - Error title
   * @param {String} message - Error message
   */
  showErrorNotification(title, message) {
    SpreadsheetApp.getUi().alert(`${title}: ${message}`);
  },
  
  /**
   * Shows an information alert
   * @param {String} title - Alert title
   * @param {String} message - Alert message
   */
  showInfoAlert(title, message) {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};

// ============================================================================
// LOCALE UTILITIES
// ============================================================================

/**
 * Utilities for locale-specific operations
 */
const LocaleUtil = {
  /**
   * Formats a date according to the configured locale
   * @param {Date} date - The date to format
   * @return {String} Formatted date string
   */
  formatDate(date) {
    return Utilities.formatDate(
      date, 
      Session.getScriptTimeZone(), 
      FINANCE_OVERVIEW_CONFIG.LOCALE.DATE_FORMAT
    );
  },
  
  /**
   * Gets the month name for a given month index
   * @param {Number} monthIndex - Month index (0-based)
   * @return {String} Localized month name
   */
  getMonthName(monthIndex) {
    const date = new Date(2024, monthIndex, 1);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM");
  },
  
  /**
   * Formats a number as currency
   * @param {SpreadsheetApp.Range} range - The range to format
   */
  formatAsCurrency(range) {
    const { CURRENCY_SYMBOL, CURRENCY_LOCALE } = FINANCE_OVERVIEW_CONFIG.LOCALE;
    range.setNumberFormat(`_-[$${CURRENCY_SYMBOL}-${CURRENCY_LOCALE}]\\ * #,##0.00_-;\\-[$${CURRENCY_SYMBOL}-${CURRENCY_LOCALE}]\\ * #,##0.00_-;_-[$${CURRENCY_SYMBOL}-${CURRENCY_LOCALE}]\\ * "-"??_-;_-@`);
  },
  
  /**
   * Formats a number as percentage
   * @param {SpreadsheetApp.Range} range - The range to format
   * @param {Number} decimalPlaces - Number of decimal places
   */
  formatAsPercentage(range, decimalPlaces = 1) {
    range.setNumberFormat(`0.${"0".repeat(decimalPlaces)}%`);
  }
};

// ============================================================================
// MAIN ENTRY POINT - BUILDER PATTERN
// ============================================================================

/**
 * Creates a financial overview sheet based on transaction data
 * This function will generate a complete overview sheet with dynamic categories
 * and optional sub-category display based on user preference
 */
function createFinancialOverview() {
  try {
    UIUtil.showLoadingSpinner("Generating financial overview...");
    CacheUtil.invalidateAll(); // Clear cache to ensure fresh data
    
    // Use the builder pattern for a cleaner implementation
    const builder = new FinancialOverviewBuilder();
    
    const result = builder
      .initialize()
      .processData()
      .setupHeader()
      .generateContent()
      .addNetCalculations()
      .addMetrics()
      .formatSheet()
      .applyPreferences()
      .build();
    
    UIUtil.hideLoadingSpinner();
    UIUtil.showSuccessNotification("Financial overview generated successfully!");
    
    return result;
  } catch (error) {
    UIUtil.hideLoadingSpinner();
    
    // Process the error
    if (error instanceof FinanceOverviewError) {
      error.logToSheet();
      UIUtil.showErrorNotification("Error generating overview", error.message);
    } else {
      // Convert to our custom error format and log
      const wrappedError = new FinanceOverviewError(
        "Failed to generate financial overview", 
        { originalError: error.message, stack: error.stack, severity: "high" }
      );
      wrappedError.logToSheet();
      UIUtil.showErrorNotification("Error generating overview", error.message);
    }
    
    throw error;
  }
}

/**
 * Builder class for generating the financial overview
 * @class
 */
class FinancialOverviewBuilder {
  /**
   * Initializes a new FinancialOverviewBuilder
   */
  constructor() {
    this.spreadsheet = null;
    this.overviewSheet = null;
    this.transactionSheet = null;
    this.showSubCategories = true;
    this.transactionData = null;
    this.columnIndices = null;
    this.categoryCombinations = null;
    this.groupedCombinations = null;
    this.lastContentRowIndex = 0;
  }
  
  /**
   * Initializes the builder with required sheets and data
   * @return {FinancialOverviewBuilder} This builder instance for chaining
   */
  initialize() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get or create the Overview sheet and clear it
    this.overviewSheet = getOrCreateSheet(
      this.spreadsheet, 
      FINANCE_OVERVIEW_CONFIG.SHEETS.OVERVIEW
    );
    clearSheetContent(this.overviewSheet);
    
    // Get transaction sheet
    this.transactionSheet = this.spreadsheet.getSheetByName(
      FINANCE_OVERVIEW_CONFIG.SHEETS.TRANSACTIONS
    );
    
    if (!this.transactionSheet) {
      throw new FinanceOverviewError(
        `Required sheet "${FINANCE_OVERVIEW_CONFIG.SHEETS.TRANSACTIONS}" not found`,
        { severity: "high" }
      );
    }
    
    // Get user preference for showing sub-categories
    this.showSubCategories = getUserPreference("ShowSubCategories", true);
    
    return this;
  }
  
  /**
   * Processes transaction data
   * @return {FinancialOverviewBuilder} This builder instance for chaining
   */
  processData() {
    // Get and process transaction data
    const { data, indices } = getProcessedTransactionData(this.transactionSheet);
    this.transactionData = data;
    this.columnIndices = indices;
    
    // Get category combinations with caching
    this.categoryCombinations = CacheUtil.getCachedOrCompute(
      FINANCE_OVERVIEW_CONFIG.CACHE.KEYS.CATEGORY_COMBINATIONS,
      () => getUniqueCategoryCombinations(
        this.transactionData, 
        this.columnIndices.type, 
        this.columnIndices.category, 
        this.columnIndices.subcategory, 
        this.showSubCategories
      )
    );
    
    // Group categories by type
    this.groupedCombinations = CacheUtil.getCachedOrCompute(
      FINANCE_OVERVIEW_CONFIG.CACHE.KEYS.GROUPED_COMBINATIONS,
      () => groupCategoryCombinations(this.categoryCombinations)
    );
    
    return this;
  }
  
  /**
   * Sets up the header row
   * @return {FinancialOverviewBuilder} This builder instance for chaining
   */
  setupHeader() {
    setupHeaderRow(this.overviewSheet, this.showSubCategories);
    return this;
  }
  
  /**
   * Generates the main content of the overview
   * @return {FinancialOverviewBuilder} This builder instance for chaining
   */
  generateContent() {
    let rowIndex = 2; // Start after header
    
    // Process each type in the defined order
    FINANCE_OVERVIEW_CONFIG.TYPE_ORDER.forEach(type => {
      // Skip if this type doesn't exist in the data
      if (!this.groupedCombinations[type]) return;
      
      // Add type header row
      rowIndex = addTypeHeaderRow(this.overviewSheet, type, rowIndex);
      
      // Add rows for each category/subcategory in this type
      rowIndex = addCategoryRows(
        this.overviewSheet, 
        this.groupedCombinations[type], 
        rowIndex, 
        type, 
        this.columnIndices
      );
      
      // Add subtotal for this type
      rowIndex = addTypeSubtotalRow(
        this.overviewSheet, 
        type, 
        rowIndex, 
        this.groupedCombinations[type].length
      );
      
      rowIndex += 2; // Add space between categories
    });
    
    this.lastContentRowIndex = rowIndex;
    return this;
  }
  
  /**
   * Adds net calculations to the overview
   * @return {FinancialOverviewBuilder} This builder instance for chaining
   */
  addNetCalculations() {
    this.lastContentRowIndex = addNetCalculations(
      this.overviewSheet, 
      this.lastContentRowIndex
    );
    return this;
  }
  
  /**
   * Adds metrics section to the overview
   * @return {FinancialOverviewBuilder} This builder instance for chaining
   */
  addMetrics() {
    // Create a combined config object for the FinancialAnalysisService
    const analysisConfig = {
      ...FINANCE_OVERVIEW_CONFIG,
      // Add any additional config needed by FinancialAnalysisService
      TARGET_RATES: {
        ...FINANCE_OVERVIEW_CONFIG.TARGET_RATES,
        WANTS_PLEASURE: FINANCE_OVERVIEW_CONFIG.TARGET_RATES.WANTS // Map WANTS to WANTS_PLEASURE for compatibility
      }
    };
    
    // Create and use the FinancialAnalysisService
    const analysisService = new FinancialAnalysisService(
      this.spreadsheet, 
      this.overviewSheet, 
      analysisConfig
    );
    analysisService.initialize();
    analysisService.analyze();
    
    return this;
  }
  
  /**
   * Formats the overview sheet
   * @return {FinancialOverviewBuilder} This builder instance for chaining
   */
  formatSheet() {
    formatOverviewSheet(this.overviewSheet);
    return this;
  }
  
  /**
   * Applies user preferences
   * @return {FinancialOverviewBuilder} This builder instance for chaining
   */
  applyPreferences() {
    // Show/hide sub-categories based on preference
    if (this.showSubCategories) {
      this.overviewSheet.showColumns(3, 1);
    } else {
      this.overviewSheet.hideColumns(3, 1);
    }
    return this;
  }
  
  /**
   * Finalizes the build process
   * @return {Object} Object containing information about the built overview
   */
  build() {
    return {
      sheet: this.overviewSheet,
      lastRow: this.overviewSheet.getLastRow(),
      success: true
    };
  }
}

// ============================================================================
// DATA PROCESSING
// ============================================================================

/**
 * Gets and processes transaction data from the sheet
 * @param {SpreadsheetApp.Sheet} sheet - The transaction sheet
 * @return {Object} Object containing processed data and column indices
 */
function getProcessedTransactionData(sheet) {
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData[0];
  
  // Find column indices
  const indices = {
    type: headers.indexOf("Type"),
    category: headers.indexOf("Category"),
    subcategory: headers.indexOf("Sub-Category"),
    date: headers.indexOf("Date"),
    amount: headers.indexOf("Amount"),
    shared: headers.indexOf("Shared")
  };
  
  // Validate required columns exist
  const requiredColumns = ["type", "category", "subcategory", "date", "amount"];
  const missingColumns = requiredColumns.filter(col => indices[col] < 0);
  
  if (missingColumns.length > 0) {
    throw new FinanceOverviewError(
      `Required columns not found: ${missingColumns.join(", ")}`,
      { severity: "high", headers }
    );
  }
  
  return { data: rawData, indices };
}

/**
 * Gets unique combinations of Type/Category/Subcategory from transaction data
 * @param {Array} data - Transaction data
 * @param {Number} typeCol - Column index for transaction type
 * @param {Number} categoryCol - Column index for category
 * @param {Number} subcategoryCol - Column index for subcategory
 * @param {Boolean} showSubCategories - Whether to show subcategories
 * @return {Array} List of unique category combinations
 */
function getUniqueCategoryCombinations(data, typeCol, categoryCol, subcategoryCol, showSubCategories) {
  // Use a Set to track unique combinations
  const seen = new Set();
  
  // Process all rows except header (more efficient than filter + map)
  return data.slice(1)
    // Skip empty rows with reduce instead of multiple filter calls
    .reduce((combinations, row) => {
      const type = row[typeCol];
      const category = row[categoryCol];
      const subcategory = showSubCategories ? row[subcategoryCol] : "";
      
      if (!type || !category) return combinations;
      
      const key = `${type}|${category}|${subcategory || ""}`;
      
      if (!seen.has(key)) {
        seen.add(key);
        combinations.push({
          type: type,
          category: category,
          subcategory: subcategory || ""
        });
      }
      
      return combinations;
    }, []);
}

/**
 * Groups category combinations by type and sorts appropriately
 * @param {Array} combinations - Array of category combinations
 * @return {Object} Grouped and sorted combinations
 */
function groupCategoryCombinations(combinations) {
  // Group by type using reduce
  const grouped = combinations.reduce((acc, combo) => {
    if (!acc[combo.type]) {
      acc[combo.type] = [];
    }
    acc[combo.type].push(combo);
    return acc;
  }, {});
  
  // Sort each group
  Object.keys(grouped).forEach(type => {
    grouped[type].sort((a, b) => {
      // Primary sort by category
      const categoryCompare = a.category.localeCompare(b.category);
      // Secondary sort by subcategory if categories are the same
      return categoryCompare !== 0 ? categoryCompare : 
             (a.subcategory || "").localeCompare(b.subcategory || "");
    });
  });
  
  return grouped;
}

/**
 * Builds a formula to sum transactions for a specific month
 * @param {Object} params - Parameters for the formula
 * @return {String} SUMIFS formula for the specified criteria
 */
function buildMonthlySumFormula(params) {
  const {
    type, category, subcategory, monthDate, sheetName,
    typeCol, categoryCol, subcategoryCol, dateCol, amountCol, sharedCol
  } = params;
  
  const month = monthDate.getMonth() + 1; // 1-indexed month
  const year = monthDate.getFullYear();
  
  // Calculate date range for the month
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0); // Last day of month
  
  // Format dates
  const startDateFormatted = LocaleUtil.formatDate(startDate);
  const endDateFormatted = LocaleUtil.formatDate(endDate);
  
  // Sum range
  const sumRange = `${sheetName}!${columnToLetter(amountCol)}:${columnToLetter(amountCol)}`;
  
  // Base criteria for all formulas
  const baseCriteria = [
    `${sheetName}!${columnToLetter(typeCol)}:${columnToLetter(typeCol)}, "${type}"`,
    `${sheetName}!${columnToLetter(categoryCol)}:${columnToLetter(categoryCol)}, "${category}"`,
    `${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, ">=${startDateFormatted}"`,
    `${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, "<=${endDateFormatted}"`
  ];
  
  // Add subcategory criteria if it exists
  if (subcategory) {
    baseCriteria.push(`${sheetName}!${columnToLetter(subcategoryCol)}:${columnToLetter(subcategoryCol)}, "${subcategory}"`);
  }
  
  // For expense types, handle shared expenses differently
  if (FINANCE_OVERVIEW_CONFIG.EXPENSE_TYPES.includes(type) && sharedCol > 0) {
    // Non-shared expenses (Shared = "")
    const nonSharedCriteria = [...baseCriteria];
    nonSharedCriteria.push(`${sheetName}!${columnToLetter(sharedCol)}:${columnToLetter(sharedCol)}, ""`);
    const nonSharedFormula = `SUMIFS(${sumRange}, ${nonSharedCriteria.join(", ")})`;
    
    // Shared expenses (Shared = TRUE, divided by 2)
    const sharedCriteria = [...baseCriteria];
    sharedCriteria.push(`${sheetName}!${columnToLetter(sharedCol)}:${columnToLetter(sharedCol)}, "true"`);
    const sharedFormula = `SUMIFS(${sumRange}, ${sharedCriteria.join(", ")})/2`;
    
    return `${nonSharedFormula} + (${sharedFormula})`;
  }
  
  // Standard formula for non-shared items
  return `SUMIFS(${sumRange}, ${baseCriteria.join(", ")})`;
}

// ============================================================================
// UI GENERATION
// ============================================================================

/**
 * Clears all content and formatting from a sheet
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to clear
 */
function clearSheetContent(sheet) {
  sheet.clear(); // Clear existing content
  sheet.clearFormats(); // Clear existing formats
  // Clear check boxes
  sheet.getRange("A1:Z1000").setDataValidation(null);
}

/**
 * Sets up the header row in the overview sheet
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 * @param {Boolean} showSubCategories - Whether subcategories are shown
 */
function setupHeaderRow(sheet, showSubCategories) {
  // Set header values (using batch operation)
  sheet.getRange(1, 1, 1, FINANCE_OVERVIEW_CONFIG.HEADERS.length)
    .setValues([FINANCE_OVERVIEW_CONFIG.HEADERS]);
  
  // Format header row
  sheet.getRange(1, 1, 1, FINANCE_OVERVIEW_CONFIG.HEADERS.length)
    .setBackground(FINANCE_OVERVIEW_CONFIG.COLORS.UI.HEADER_BG)
    .setFontWeight("bold")
    .setFontColor(FINANCE_OVERVIEW_CONFIG.COLORS.UI.HEADER_FONT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  // Add the sub-category toggle
  const { SUBCATEGORY_TOGGLE } = FINANCE_OVERVIEW_CONFIG.UI;
  sheet.getRange(SUBCATEGORY_TOGGLE.LABEL_CELL)
    .setValue(SUBCATEGORY_TOGGLE.LABEL_TEXT)
    .setFontWeight("bold");
  
  // Add the checkbox
  const checkbox = sheet.getRange(SUBCATEGORY_TOGGLE.CHECKBOX_CELL);
  checkbox.insertCheckboxes();
  checkbox.setValue(showSubCategories);
  checkbox.setNote(SUBCATEGORY_TOGGLE.NOTE_TEXT);
  
  // Freeze the header row
  sheet.setFrozenRows(1);
  
  // Set column width for the Shared? column
  sheet.setColumnWidth(4, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.SHARED);
}

/**
 * Adds a type header row to the overview sheet
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 * @param {String} type - The transaction type
 * @param {Number} rowIndex - The current row index
 * @return {Number} The next row index
 */
function addTypeHeaderRow(sheet, type, rowIndex) {
  // Get colors for this type
  const typeColors = getTypeColors(type);
  
  // Add Type header row with appropriate color
  sheet.getRange(rowIndex, 1).setValue(type);
  sheet.getRange(rowIndex, 1, 1, FINANCE_OVERVIEW_CONFIG.HEADERS.length)
    .setBackground(typeColors.BG)
    .setFontWeight("bold")
    .setFontColor(typeColors.FONT);
  
  return rowIndex + 1;
}

/**
 * Adds category and subcategory rows to the overview sheet
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 * @param {Array} combinations - Array of category combinations for this type
 * @param {Number} rowIndex - The current row index
 * @param {String} type - The transaction type
 * @param {Object} columnIndices - Column indices for transaction data
 * @return {Number} The next row index
 */
function addCategoryRows(sheet, combinations, rowIndex, type, columnIndices) {
  if (combinations.length === 0) return rowIndex;
  
  // Batch operation for setting values
  const startRow = rowIndex;
  const numRows = combinations.length;
  
  // Create arrays for bulk updates
  const categoryData = Array(numRows).fill().map(() => Array(3).fill(""));
  
  // Prepare data for bulk insert
  combinations.forEach((combo, index) => {
    categoryData[index][0] = combo.type;
    categoryData[index][1] = combo.category;
    categoryData[index][2] = combo.subcategory;
  });
  
  // Set all values at once
  sheet.getRange(startRow, 1, numRows, 3).setValues(categoryData);
  
  // Set checkboxes for Shared column for expense types
  if (FINANCE_OVERVIEW_CONFIG.EXPENSE_TYPES.includes(type)) {
    sheet.getRange(startRow, 4, numRows, 1).insertCheckboxes();
  }
  
  // Batch prepare all formulas
  const monthFormulas = [];
  
  // For each category combination...
  for (let i = 0; i < combinations.length; i++) {
    const combo = combinations[i];
    const currentRow = startRow + i;
    
    // Prepare formulas for all months (columns 5-16)
    const rowFormulas = [];
    
    for (let monthCol = 5; monthCol <= 16; monthCol++) {
      const monthDate = new Date(2024, monthCol - 5, 1); // Adjusted for month column offset
      
      const formulaParams = {
        type: combo.type,
        category: combo.category,
        subcategory: combo.subcategory,
        monthDate: monthDate,
        sheetName: FINANCE_OVERVIEW_CONFIG.SHEETS.TRANSACTIONS,
        typeCol: columnIndices.type + 1,
        categoryCol: columnIndices.category + 1,
        subcategoryCol: columnIndices.subcategory + 1,
        dateCol: columnIndices.date + 1,
        amountCol: columnIndices.amount + 1,
        sharedCol: columnIndices.shared + 1
      };
      
      rowFormulas.push(buildMonthlySumFormula(formulaParams));
    }
    
    // Add formula for each month column
    if (FINANCE_OVERVIEW_CONFIG.PERFORMANCE.USE_BATCH_OPERATIONS) {
      // Apply in batches for better performance
      sheet.getRange(currentRow, 5, 1, 12).setFormulas([rowFormulas]);
    } else {
      // Apply individually if batch operations disabled
      for (let monthCol = 0; monthCol < 12; monthCol++) {
        sheet.getRange(currentRow, monthCol + 5).setFormula(rowFormulas[monthCol]);
      }
    }
    
    // Add total formula in column 17
    sheet.getRange(currentRow, 17)
      .setFormula(`=SUM(E${currentRow}:P${currentRow})`);
    
    // Add average formula in column 18
    sheet.getRange(currentRow, 18)
      .setFormula(`=AVERAGE(E${currentRow}:P${currentRow})`);
    
    // Add styling for subcategories and main categories
    if (combo.subcategory) {
      // Indent subcategories
      sheet.getRange(currentRow, 3).setIndent(5);
    } else {
      // Bold main categories
      sheet.getRange(currentRow, 2).setFontWeight("bold");
    }
  }
  
  // Apply conditional formatting for all rows at once
  const valueRange = sheet.getRange(startRow, 5, numRows, 13); // Columns E through Q
  
  // Format all values as currency
  formatRangeAsCurrency(valueRange);
  
  // Color income/expense values
  for (let i = 0; i < numRows; i++) {
    for (let col = 5; col <= 17; col++) {
      const cell = sheet.getRange(startRow + i, col);
      
      if (type === FINANCE_OVERVIEW_CONFIG.TRANSACTION_TYPES.INCOME) {
        cell.setFontColor(FINANCE_OVERVIEW_CONFIG.COLORS.UI.INCOME_FONT);
      } else if (FINANCE_OVERVIEW_CONFIG.EXPENSE_TYPES.includes(type)) {
        cell.setFontColor(FINANCE_OVERVIEW_CONFIG.COLORS.UI.EXPENSE_FONT);
      }
    }
  }
  
  return rowIndex + numRows;
}

/**
 * Adds a subtotal row for a transaction type
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 * @param {String} type - The transaction type
 * @param {Number} rowIndex - The current row index
 * @param {Number} rowCount - The number of rows for this type
 * @return {Number} The next row index
 */
function addTypeSubtotalRow(sheet, type, rowIndex, rowCount) {
  // Get colors for this type
  const typeColors = getTypeColors(type);
  
  // Add subtotal for this type
  sheet.getRange(rowIndex, 1).setValue(`Total ${type}`);
  sheet.getRange(rowIndex, 1, 1, FINANCE_OVERVIEW_CONFIG.HEADERS.length)
    .setBackground(typeColors.BG)
    .setFontWeight("bold")
    .setFontColor(typeColors.FONT);
  
  // Add subtotal formulas for each month column using batch operations
  const formulas = [];
  const startRow = rowIndex - rowCount;
  const endRow = rowIndex - 1;
  
  for (let monthCol = 5; monthCol <= 17; monthCol++) {
    formulas.push(`=SUM(${columnToLetter(monthCol)}${startRow}:${columnToLetter(monthCol)}${endRow})`);
  }
  
  // Set all formulas at once for better performance
  sheet.getRange(rowIndex, 5, 1, 13).setFormulas([formulas]);
  
  return rowIndex + 1;
}

/**
 * Adds net calculation rows to the overview sheet
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 * @param {Number} startRow - The row to start adding net calculations
 * @return {Number} The next row index
 */
function addNetCalculations(sheet, startRow) {
  // Find rows containing total Income, Expenses, and Savings
  const data = sheet.getDataRange().getValues();
  
  // Find the rows containing the totals we need
  const totals = findTotalRows(data);
  const { incomeRow, expensesRow, savingsRow } = totals;
  
  if (!incomeRow || !expensesRow) {
    // Couldn't find required rows
    return startRow;
  }
  
  // Add a section header for Net Calculations
  sheet.getRange(startRow, 1).setValue("Net Calculations");
  sheet.getRange(startRow, 1, 1, FINANCE_OVERVIEW_CONFIG.HEADERS.length)
    .setBackground(FINANCE_OVERVIEW_CONFIG.COLORS.UI.NET_BG)
    .setFontWeight("bold")
    .setFontColor(FINANCE_OVERVIEW_CONFIG.COLORS.UI.NET_FONT);
  
  startRow++;
  
  // Add Net (Income - Expenses) row
  sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses)");
  sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
  
  // Create all formulas at once
  const netFormulas = [];
  for (let col = 5; col <= 17; col++) {
    netFormulas.push(`=${columnToLetter(col)}${incomeRow}-${columnToLetter(col)}${expensesRow}`);
  }
  
  // Apply formulas in batch
  sheet.getRange(startRow, 5, 1, 13).setFormulas([netFormulas]);
  
  // Format the cells
  formatRangeAsCurrency(sheet.getRange(startRow, 5, 1, 13));
  
  startRow++;
  
  // Add additional calculations if savings data exists
  if (savingsRow) {
    // Add Total Expenses + Savings row
    sheet.getRange(startRow, 1).setValue("Total Expenses + Savings");
    sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
    
    // Create formulas
    const expSavFormulas = [];
    for (let col = 5; col <= 17; col++) {
      expSavFormulas.push(`=${columnToLetter(col)}${expensesRow}+${columnToLetter(col)}${savingsRow}`);
    }
    
    // Apply formulas in batch
    sheet.getRange(startRow, 5, 1, 13).setFormulas([expSavFormulas]);
    formatRangeAsCurrency(sheet.getRange(startRow, 5, 1, 13));
    
    startRow++;
    
    // Add Net (Income - Expenses - Savings) row
    sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses - Savings)");
    sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
    
    // Create formulas
    const totalNetFormulas = [];
    for (let col = 5; col <= 17; col++) {
      totalNetFormulas.push(`=${columnToLetter(col)}${incomeRow}-${columnToLetter(col)}${expensesRow}-${columnToLetter(col)}${savingsRow}`);
    }
    
    // Apply formulas in batch
    sheet.getRange(startRow, 5, 1, 13).setFormulas([totalNetFormulas]);
    formatRangeAsCurrency(sheet.getRange(startRow, 5, 1, 13));
  }
  
  // Add a bottom border to the last row
  sheet.getRange(startRow, 1, 1, FINANCE_OVERVIEW_CONFIG.HEADERS.length).setBorder(
    null, null, true, null, null, null, 
    FINANCE_OVERVIEW_CONFIG.COLORS.UI.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  return startRow + 2; // Add space after net calculations
}

/**
 * Finds the row numbers for total values in the data
 * @param {Array} data - The sheet data
 * @return {Object} Object containing row indices for totals
 */
function findTotalRows(data) {
  const totals = {
    incomeRow: null,
    expensesRow: null,
    savingsRow: null
  };
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "Total Income") totals.incomeRow = i + 1;
    if (data[i][0] === "Total Expenses") totals.expensesRow = i + 1;
    if (data[i][0] === "Total Savings") totals.savingsRow = i + 1;
  }
  
  return totals;
}

/**
 * Adds key metrics section to the overview sheet
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 * @param {Number} startRow - The row to start adding key metrics
 */
function addKeyMetricsSection(sheet, startRow) {
  // Find the rows containing the totals we need
  const data = sheet.getDataRange().getValues();
  const { incomeRow, expensesRow, savingsRow } = findTotalRows(data);
  
  if (!incomeRow) return; // Can't proceed without income data
  
  // Add Key Metrics header
  sheet.getRange(startRow, 10).setValue("Key Metrics");
  sheet.getRange(startRow, 10, 1, 3)
    .setBackground(FINANCE_OVERVIEW_CONFIG.COLORS.UI.HEADER_BG)
    .setFontWeight("bold")
    .setFontColor(FINANCE_OVERVIEW_CONFIG.COLORS.UI.HEADER_FONT)
    .setHorizontalAlignment("center");
  
  startRow++;
  
  // Create metrics table header
  sheet.getRange(startRow, 10, 1, 3)
    .setValues([["Metric", "Value", "Target"]])
    .setBackground("#F5F5F5")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  startRow++;
  
  // Create arrays to hold metrics data
  const metricsData = [];
  
  // Add Savings Rate if we have savings data
  if (savingsRow) {
    metricsData.push({
      name: "Savings Rate",
      valueFormula: `=Q${savingsRow}/Q${incomeRow}`, // Use Total column (Q)
      target: 0.2,
      compareFunction: (value, target) => value < target,
      severity: "medium"
    });
  }
  
  // Add Expenses/Income Ratio if we have expense data
  if (expensesRow) {
    metricsData.push({
      name: "Expenses/Income Ratio",
      valueFormula: `=Q${expensesRow}/Q${incomeRow}`, // Use Total column (Q)
      target: 0.8,
      compareFunction: (value, target) => value > target,
      severity: "medium"
    });
  }
  
  // Set values and formulas for metrics
  metricsData.forEach((metric, index) => {
    const currentRow = startRow + index;
    
    // Set metric name and target
    sheet.getRange(currentRow, 10).setValue(metric.name);
    sheet.getRange(currentRow, 12).setValue(metric.target);
    
    // Set formula for value
    sheet.getRange(currentRow, 11).setFormula(metric.valueFormula);
    
    // Apply styling
    if (index % 2 === 0) {
      sheet.getRange(currentRow, 10, 1, 3).setBackground(FINANCE_OVERVIEW_CONFIG.COLORS.UI.METRICS_BG);
    }
  });
  
  // Format as percentage
  if (metricsData.length > 0) {
    sheet.getRange(startRow, 11, metricsData.length, 2).setNumberFormat("0.0%");
    
    // Add conditional formatting for metrics
    const rules = sheet.getConditionalFormatRules();
    
    metricsData.forEach((metric, index) => {
      const currentRow = startRow + index;
      const targetCell = sheet.getRange(currentRow, 12);
      const valueCell = sheet.getRange(currentRow, 11);
      
      // Create appropriate conditional format rule
      let rule;
      if (metric.compareFunction(0.5, 0.3)) { // Test if this is a "less than" comparison
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenCellNotEmpty()
          .setFormula(`K${currentRow}<M${currentRow}`)
          .setBackground("#FFCDD2") // Light red when below target
          .setRanges([valueCell])
          .build();
      } else {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenCellNotEmpty()
          .setFormula(`K${currentRow}>M${currentRow}`)
          .setBackground("#FFCDD2") // Light red when above target
          .setRanges([valueCell])
          .build();
      }
      
      rules.push(rule);
    });
    
    sheet.setConditionalFormatRules(rules);
    
    // Add a border around the metrics table
    sheet.getRange(startRow - 1, 10, metricsData.length + 1, 3).setBorder(
      true, true, true, true, true, true, 
      "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID
    );
  }
  
  // Add Expense Categories section
  const expenseStartRow = startRow + metricsData.length + 2;
  
  addExpenseCategoriesSection(sheet, expenseStartRow, data, incomeRow, expensesRow);
}

/**
 * Adds the expense categories section to the overview sheet
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 * @param {Number} startRow - The row to start adding expense categories
 * @param {Array} data - The sheet data
 * @param {Number} incomeRow - The row containing total income
 * @param {Number} expensesRow - The row containing total expenses
 */
function addExpenseCategoriesSection(sheet, startRow, data, incomeRow, expensesRow) {
  // Add Expense Categories header
  sheet.getRange(startRow, 10).setValue("Expense Categories");
  sheet.getRange(startRow, 10, 1, 6)
    .setBackground(FINANCE_OVERVIEW_CONFIG.COLORS.UI.HEADER_BG)
    .setFontWeight("bold")
    .setFontColor(FINANCE_OVERVIEW_CONFIG.COLORS.UI.HEADER_FONT)
    .setHorizontalAlignment("center");
  
  startRow++;
  
  // Add table headers
  sheet.getRange(startRow, 10, 1, 6).setValues([
    ["Expense", "Amount", "Rate", "Target Rate", "% change", "Amount change"]
  ]);
  
  sheet.getRange(startRow, 10, 1, 6)
    .setBackground("#F5F5F5")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  startRow++;
  
  // Find expense categories
  const expenseCategories = findExpenseCategories(data);
  
  // Add rows for each expense category in batch
  if (expenseCategories.length > 0) {
    // Prepare data arrays
    const categoryNames = [];
    const amountFormulas = [];
    const rateFormulas = [];
    const targetRates = [];
    const changeFormulas = [];
    const amountChangeFormulas = [];
    
    expenseCategories.forEach((category, index) => {
      const currentRow = startRow + index;
      
      // Set values
      categoryNames.push([category.category]);
      
      // Set formulas
      amountFormulas.push([`=Q${category.row}`]); // Amount (using Total column)
      rateFormulas.push([`=IFERROR(K${currentRow}/Q${incomeRow}, 0)`]); // Rate
      
      // Set target rate based on expense type
      let targetRate = FINANCE_OVERVIEW_CONFIG.TARGET_RATES.DEFAULT;
      if (category.type === "Essentials") {
        targetRate = FINANCE_OVERVIEW_CONFIG.TARGET_RATES.ESSENTIALS;
      } else if (category.type === "Wants/Pleasure") {
        targetRate = FINANCE_OVERVIEW_CONFIG.TARGET_RATES.WANTS;
      } else if (category.type === "Extra") {
        targetRate = FINANCE_OVERVIEW_CONFIG.TARGET_RATES.EXTRA;
      }
      targetRates.push([targetRate]);
      
      // Comparison formulas
      changeFormulas.push([`=IFERROR((L${currentRow}-M${currentRow})/M${currentRow}, 0)`]);
      amountChangeFormulas.push([`=IFERROR(K${currentRow}-(Q${incomeRow}*M${currentRow}), 0)`]);
    });
    
    // Apply values and formulas in batches
    sheet.getRange(startRow, 10, expenseCategories.length, 1).setValues(categoryNames);
    sheet.getRange(startRow, 11, expenseCategories.length, 1).setFormulas(amountFormulas);
    sheet.getRange(startRow, 12, expenseCategories.length, 1).setFormulas(rateFormulas);
    sheet.getRange(startRow, 13, expenseCategories.length, 1).setValues(targetRates);
    sheet.getRange(startRow, 14, expenseCategories.length, 1).setFormulas(changeFormulas);
    sheet.getRange(startRow, 15, expenseCategories.length, 1).setFormulas(amountChangeFormulas);
    
    // Apply formatting in batches
    sheet.getRange(startRow, 11, expenseCategories.length, 1).setNumberFormat(getCurrencyFormat());
    sheet.getRange(startRow, 12, expenseCategories.length, 1).setNumberFormat("0.0%");
    sheet.getRange(startRow, 13, expenseCategories.length, 1).setNumberFormat("0.0%");
    sheet.getRange(startRow, 14, expenseCategories.length, 1).setNumberFormat("0.0%");
    sheet.getRange(startRow, 15, expenseCategories.length, 1).setNumberFormat(getCurrencyFormat());
    
    // Apply alternating row colors
    for (let i = 0; i < expenseCategories.length; i++) {
      if (i % 2 === 0) {
        sheet.getRange(startRow + i, 10, 1, 6).setBackground(FINANCE_OVERVIEW_CONFIG.COLORS.UI.METRICS_BG);
      } else {
        sheet.getRange(startRow + i, 10, 1, 6).setBackground("#F5F5F5");
      }
    }
  }
  
  // Add Total Expenses row with distinct formatting
  const totalRow = startRow + expenseCategories.length;
  sheet.getRange(totalRow, 10).setValue("Total Expenses");
  sheet.getRange(totalRow, 11).setFormula(`=Q${expensesRow}`); // Use Total column (Q)
  sheet.getRange(totalRow, 12).setFormula(`=IFERROR(K${totalRow}/Q${incomeRow}, 0)`);
  sheet.getRange(totalRow, 13).setValue(1); // Target 100%
  sheet.getRange(totalRow, 14).setFormula(`=IFERROR((L${totalRow}-M${totalRow})/M${totalRow}, 0)`);
  sheet.getRange(totalRow, 15).setFormula(`=IFERROR(K${totalRow}-(Q${incomeRow}*M${totalRow}), 0)`);
  
  // Format total row
  sheet.getRange(totalRow, 10, 1, 6)
    .setBackground(FINANCE_OVERVIEW_CONFIG.COLORS.UI.HEADER_BG)
    .setFontWeight("bold")
    .setFontColor(FINANCE_OVERVIEW_CONFIG.COLORS.UI.HEADER_FONT);
  
  sheet.getRange(totalRow, 11).setNumberFormat(getCurrencyFormat());
  sheet.getRange(totalRow, 12, 1, 2).setNumberFormat("0.0%");
  sheet.getRange(totalRow, 14).setNumberFormat("0.0%");
  sheet.getRange(totalRow, 15).setNumberFormat(getCurrencyFormat());
  
  // Add borders to the expense table
  sheet.getRange(startRow - 1, 10, expenseCategories.length + 2, 6).setBorder(
    true, true, true, true, true, true, 
    "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Create expenditure charts
  if (expenseCategories.length > 0) {
    createExpenditureCharts(sheet, startRow, totalRow - 1, 10);
  }
}

/**
 * Finds expense categories from the sheet data
 * @param {Array} data - The sheet data
 * @return {Array} Array of expense category objects
 */
function findExpenseCategories(data) {
  return data.reduce((categories, row, index) => {
    // Check if this row has a type that's considered an expense and has a category
    if (FINANCE_OVERVIEW_CONFIG.EXPENSE_TYPES.includes(row[0]) && row[1]) {
      categories.push({
        category: row[1],
        type: row[0],
        row: index + 1
      });
    }
    return categories;
  }, []);
}

/**
 * Creates charts for expenditure breakdown
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 * @param {Number} startRow - The start row for chart data
 * @param {Number} endRow - The end row for chart data
 * @param {Number} categoryCol - The column containing category names
 */
function createExpenditureCharts(sheet, startRow, endRow, categoryCol) {
  // Define chart data range (category name and amount)
  const dataRange = sheet.getRange(startRow, categoryCol, endRow - startRow + 1, 2);
  
  // Create a pie chart with enhanced styling
  const pieChartBuilder = sheet.newChart();
  pieChartBuilder
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setPosition(startRow, categoryCol + 6, 0, 0)
    .setOption('title', 'Expenditure Breakdown')
    .setOption('titleTextStyle', {
      color: FINANCE_OVERVIEW_CONFIG.COLORS.CHART.TITLE,
      fontSize: 16,
      bold: true
    })
    .setOption('pieSliceText', 'percentage')
    .setOption('pieHole', 0.4) // Create a donut chart for more modern look
    .setOption('legend', { 
      position: 'right',
      textStyle: {
        color: FINANCE_OVERVIEW_CONFIG.COLORS.CHART.TEXT,
        fontSize: 12
      }
    })
    .setOption('colors', FINANCE_OVERVIEW_CONFIG.COLORS.CHART.SERIES)
    .setOption('width', 450)
    .setOption('height', 300);
  
  // Add the chart to the sheet
  sheet.insertChart(pieChartBuilder.build());
  
  // Create a second chart - a column chart showing expense categories vs target
  const columnChartBuilder = sheet.newChart();
  
  // Define data ranges for the column chart
  const rateRange = sheet.getRange(startRow, categoryCol, endRow - startRow, 1); // Category names
  const valueRange = sheet.getRange(startRow, categoryCol + 2, endRow - startRow, 1); // Rate column
  const targetRange = sheet.getRange(startRow, categoryCol + 3, endRow - startRow, 1); // Target Rate column
  
  columnChartBuilder
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(rateRange)
    .addRange(valueRange)
    .addRange(targetRange)
    .setPosition(startRow + 15, categoryCol + 6, 0, 0)
    .setOption('title', 'Expense Rates vs Targets')
    .setOption('titleTextStyle', {
      color: FINANCE_OVERVIEW_CONFIG.COLORS.CHART.TITLE,
      fontSize: 16,
      bold: true
    })
    .setOption('legend', { 
      position: 'top',
      textStyle: {
        color: FINANCE_OVERVIEW_CONFIG.COLORS.CHART.TEXT,
        fontSize: 12
      }
    })
    .setOption('colors', [FINANCE_OVERVIEW_CONFIG.COLORS.UI.EXPENSE_FONT, FINANCE_OVERVIEW_CONFIG.COLORS.UI.INCOME_FONT])
    .setOption('hAxis', {
      title: 'Category',
      titleTextStyle: {color: FINANCE_OVERVIEW_CONFIG.COLORS.CHART.TEXT},
      textStyle: {color: FINANCE_OVERVIEW_CONFIG.COLORS.CHART.TEXT, fontSize: 10}
    })
    .setOption('vAxis', {
      title: 'Rate (% of Income)',
      titleTextStyle: {color: FINANCE_OVERVIEW_CONFIG.COLORS.CHART.TEXT},
      textStyle: {color: FINANCE_OVERVIEW_CONFIG.COLORS.CHART.TEXT},
      format: 'percent'
    })
    .setOption('bar', {groupWidth: '75%'})
    .setOption('isStacked', false);
  
  // Add the column chart to the sheet
  sheet.insertChart(columnChartBuilder.build());
}

/**
 * Formats the overview sheet for better readability
 * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
 */
function formatOverviewSheet(sheet) {
  const lastRow = sheet.getLastRow();
  
  // Set column widths
  sheet.setColumnWidth(1, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.TYPE);
  sheet.setColumnWidth(2, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.CATEGORY);
  sheet.setColumnWidth(3, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.SUBCATEGORY);
  sheet.setColumnWidth(4, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.SHARED);
  
  // Set month column widths
  for (let i = 5; i <= 16; i++) {
    sheet.setColumnWidth(i, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.MONTH);
  }
  
  // Set Total and Average column widths
  sheet.setColumnWidth(17, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.AVERAGE); // Total column
  sheet.setColumnWidth(18, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.AVERAGE); // Average column
  
  // Set metrics section column widths
  sheet.setColumnWidth(10, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.EXPENSE_CATEGORY);
  sheet.setColumnWidth(11, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.AMOUNT);
  sheet.setColumnWidth(12, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.RATE);
  sheet.setColumnWidth(13, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.RATE);
  sheet.setColumnWidth(14, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.RATE);
  sheet.setColumnWidth(15, FINANCE_OVERVIEW_CONFIG.UI.COLUMN_WIDTHS.AMOUNT);
  
  // Add gridlines and borders to improve readability
  const data = sheet.getDataRange().getValues();
  
  // Add bottom borders to total rows
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].startsWith("Total ")) {
      sheet.getRange(i + 1, 1, 1, FINANCE_OVERVIEW_CONFIG.HEADERS.length).setBorder(
        null, null, true, null, null, null, 
        FINANCE_OVERVIEW_CONFIG.COLORS.UI.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    }
  }
}

// ============================================================================
// EVENT HANDLERS
// ============================================================================

/**
 * Handles edits to the overview sheet, specifically for the sub-category toggle checkbox
 * Must be triggered from the onEdit(e) function
 * @param {Object} e - The edit event object
 */
function handleOverviewSheetEdits(e) {
  // Check if the edit was in the Overview sheet
  if (e.range.getSheet().getName() !== FINANCE_OVERVIEW_CONFIG.SHEETS.OVERVIEW) return;
  
  // Check if the edit was to the checkbox cell (T1)
  if (e.range.getA1Notation() === FINANCE_OVERVIEW_CONFIG.UI.SUBCATEGORY_TOGGLE.CHECKBOX_CELL) {
    const newValue = e.range.getValue();
    
    // Update the user preference
    setUserPreference("ShowSubCategories", newValue);
    
    // Show loading toast
    UIUtil.showLoadingSpinner("Updating overview...");
    
    // Regenerate the overview
    try {
      createFinancialOverview();
      
      const status = newValue ? "showing" : "hiding";
      UIUtil.showSuccessNotification(`Overview updated, ${status} sub-categories`);
    } catch (error) {
      UIUtil.showErrorNotification("Update failed", error.message);
    }
  }
}

// ============================================================================
// TESTING FRAMEWORK
// ============================================================================

/**
 * Test framework for financial overview module
 */
const TEST = {
  /**
   * Runs all tests in the test framework
   * @return {Array} Array of test results
   */
  runTests: function() {
    const testResults = [];
    
    for (const testName in this) {
      if (testName.startsWith('test') && typeof this[testName] === 'function') {
        try {
          this[testName]();
          testResults.push(`✓ ${testName} passed`);
        } catch (error) {
          testResults.push(`✗ ${testName} failed: ${error.message}`);
        }
      }
    }
    
    Logger.log(testResults.join('\n'));
    return testResults;
  },
  
  assertEquals: function(expected, actual, message) {
    if (expected !== actual) {
      throw new Error(`${message || 'Assertion failed'}: expected ${expected}, got ${actual}`);
    }
  },
  
  assertContains: function(substring, fullString, message) {
    if (!fullString.includes(substring)) {
      throw new Error(`${message || 'Assertion failed'}: expected "${fullString}" to contain "${substring}"`);
    }
  },
  
  // Example tests
  testBuildMonthlySumFormula: function() {
    const params = {
      type: 'Income',
      category: 'Salary',
      subcategory: '',
      monthDate: new Date(2024, 0, 1),
      sheetName: 'Transactions',
      typeCol: 1,
      categoryCol: 2,
      subcategoryCol: 3,
      dateCol: 4,
      amountCol: 5,
      sharedCol: 6
    };
    
    const formula = buildMonthlySumFormula(params);
    this.assertContains('SUMIFS', formula, 'Formula should use SUMIFS');
    this.assertContains('"Income"', formula, 'Formula should include type');
    this.assertContains('"Salary"', formula, 'Formula should include category');
  },
  
  testGetTypeColors: function() {
    const incomeColors = getTypeColors(FINANCE_OVERVIEW_CONFIG.TRANSACTION_TYPES.INCOME);
    this.assertEquals(
      FINANCE_OVERVIEW_CONFIG.COLORS.TYPE_HEADERS.INCOME.BG,
      incomeColors.BG,
      'Should return correct background color for Income'
    );
    
    const defaultColors = getTypeColors('Unknown');
    this.assertEquals(
      FINANCE_OVERVIEW_CONFIG.COLORS.TYPE_HEADERS.DEFAULT.BG,
      defaultColors.BG,
      'Should return default colors for unknown type'
    );
  }
};

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Gets a Date object for the month represented by a column index
 * @param {Number} colIndex - The column index (1-based)
 * @return {Date} Date object for the first day of the month
 */
function getMonthDateFromColIndex(colIndex) {
  // Column 5 = Jan 2024, 6 = Feb 2024, etc. (adjusted for Shared? column)
  const monthOffset = colIndex - 5;
  return new Date(2024, monthOffset, 1); // 0 = January (0-indexed)
}

/**
 * Gets the colors for a specific transaction type
 * @param {String} type - The transaction type
 * @return {Object} Object containing background and font colors
 */
function getTypeColors(type) {
  let colors = FINANCE_OVERVIEW_CONFIG.COLORS.TYPE_HEADERS.DEFAULT;
  
  if (type === FINANCE_OVERVIEW_CONFIG.TRANSACTION_TYPES.INCOME) {
    colors = FINANCE_OVERVIEW_CONFIG.COLORS.TYPE_HEADERS.INCOME;
  } else if (type === FINANCE_OVERVIEW_CONFIG.TRANSACTION_TYPES.ESSENTIALS) {
    colors = FINANCE_OVERVIEW_CONFIG.COLORS.TYPE_HEADERS.ESSENTIALS;
  } else if (type === "Wants/Pleasure") {
    colors = FINANCE_OVERVIEW_CONFIG.COLORS.TYPE_HEADERS.WANTS_PLEASURE;
  } else if (type === FINANCE_OVERVIEW_CONFIG.TRANSACTION_TYPES.EXTRA) {
    colors = FINANCE_OVERVIEW_CONFIG.COLORS.TYPE_HEADERS.EXTRA;
  } else if (type === FINANCE_OVERVIEW_CONFIG.TRANSACTION_TYPES.SAVINGS) {
    colors = FINANCE_OVERVIEW_CONFIG.COLORS.TYPE_HEADERS.SAVINGS;
  }
  
  return colors;
}

/**
 * Formats a range of cells as currency
 * @param {SpreadsheetApp.Range} range - The range to format
 */
function formatRangeAsCurrency(range) {
  range.setNumberFormat(getCurrencyFormat());
}

/**
 * Gets the currency format string for the current locale
 * @return {String} Currency format string
 */
function getCurrencyFormat() {
  const { CURRENCY_SYMBOL, CURRENCY_LOCALE } = FINANCE_OVERVIEW_CONFIG.LOCALE;
  return `_-[$${CURRENCY_SYMBOL}-${CURRENCY_LOCALE}]\\ * #,##0.00_-;\\-[$${CURRENCY_SYMBOL}-${CURRENCY_LOCALE}]\\ * #,##0.00_-;_-[$${CURRENCY_SYMBOL}-${CURRENCY_LOCALE}]\\ * "-"??_-;_-@`;
}

// Export the necessary functions
// These functions would be exported in a typical module system,
// but Google Apps Script makes all functions global by default

// For documentation/usage instructions
/**
 * Documentation for the Financial Overview module
 * 
 * Public functions:
 * - createFinancialOverview(): Generates a complete financial overview sheet
 * - handleOverviewSheetEdits(e): Handles edits to the overview sheet (to be called from onEdit)
 * 
 * Usage example:
 * 1. Call createFinancialOverview() to generate the overview
 * 2. In your onEdit function, include: handleOverviewSheetEdits(e)
 */
