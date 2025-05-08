/**
 * Financial Planning Tools - Financial Overview Generator
 * 
 * This module creates a comprehensive financial overview sheet based on transaction data.
 * It generates a complete overview with dynamic categories and optional sub-category display
 * based on user preference.
 * 
 * Version: 2.1.0
 * Last Updated: 2025-05-08
 */

// Create the FinanceOverview module within the FinancialPlanner namespace
FinancialPlanner.FinanceOverview = (function(utils, uiService, cacheService, errorService, config) {
  // ============================================================================
  // PRIVATE IMPLEMENTATION
  // ============================================================================
  
  /**
   * Gets and processes transaction data from the sheet
   * @param {SpreadsheetApp.Sheet} sheet - The transaction sheet
   * @return {Object} Object containing processed data and column indices
   * @private
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
      throw errorService.create(
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
   * @private
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
   * @private
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
   * @private
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
    const startDateFormatted = formatDate(startDate);
    const endDateFormatted = formatDate(endDate);
    
    // Sum range
    const sumRange = `${sheetName}!${utils.columnToLetter(amountCol)}:${utils.columnToLetter(amountCol)}`;
    
    // Base criteria for all formulas
    const baseCriteria = [
      `${sheetName}!${utils.columnToLetter(typeCol)}:${utils.columnToLetter(typeCol)}, "${type}"`,
      `${sheetName}!${utils.columnToLetter(categoryCol)}:${utils.columnToLetter(categoryCol)}, "${category}"`,
      `${sheetName}!${utils.columnToLetter(dateCol)}:${utils.columnToLetter(dateCol)}, ">=${startDateFormatted}"`,
      `${sheetName}!${utils.columnToLetter(dateCol)}:${utils.columnToLetter(dateCol)}, "<=${endDateFormatted}"`
    ];
    
    // Add subcategory criteria if it exists
    if (subcategory) {
      baseCriteria.push(`${sheetName}!${utils.columnToLetter(subcategoryCol)}:${utils.columnToLetter(subcategoryCol)}, "${subcategory}"`);
    }
    
    // For expense types, handle shared expenses differently
    const expenseTypes = config.getSection('EXPENSE_TYPES');
    if (expenseTypes.includes(type) && sharedCol > 0) {
      // Non-shared expenses (Shared = "")
      const nonSharedCriteria = [...baseCriteria];
      nonSharedCriteria.push(`${sheetName}!${utils.columnToLetter(sharedCol)}:${utils.columnToLetter(sharedCol)}, ""`);
      const nonSharedFormula = `SUMIFS(${sumRange}, ${nonSharedCriteria.join(", ")})`;
      
      // Shared expenses (Shared = TRUE, divided by 2)
      const sharedCriteria = [...baseCriteria];
      sharedCriteria.push(`${sheetName}!${utils.columnToLetter(sharedCol)}:${utils.columnToLetter(sharedCol)}, "true"`);
      const sharedFormula = `SUMIFS(${sumRange}, ${sharedCriteria.join(", ")})/2`;
      
      return `${nonSharedFormula} + (${sharedFormula})`;
    }
    
    // Standard formula for non-shared items
    return `SUMIFS(${sumRange}, ${baseCriteria.join(", ")})`;
  }
  
  /**
   * Formats a date according to the configured locale
   * @param {Date} date - The date to format
   * @return {String} Formatted date string
   * @private
   */
  function formatDate(date) {
    return Utilities.formatDate(
      date, 
      Session.getScriptTimeZone(), 
      config.getSection('LOCALE').DATE_FORMAT
    );
  }
  
  /**
   * Clears all content and formatting from a sheet
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to clear
   * @private
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
   * @private
   */
  function setupHeaderRow(sheet, showSubCategories) {
    const headers = config.getSection('HEADERS');
    const uiConfig = config.getSection('UI');
    const colors = config.getSection('COLORS').UI;
    
    // Set header values (using batch operation)
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(colors.HEADER_BG)
      .setFontWeight("bold")
      .setFontColor(colors.HEADER_FONT)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    // Add the sub-category toggle
    const { SUBCATEGORY_TOGGLE } = uiConfig;
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
    sheet.setColumnWidth(4, uiConfig.COLUMN_WIDTHS.SHARED);
  }
  
  /**
   * Adds a type header row to the overview sheet
   * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
   * @param {String} type - The transaction type
   * @param {Number} rowIndex - The current row index
   * @return {Number} The next row index
   * @private
   */
  function addTypeHeaderRow(sheet, type, rowIndex) {
    // Get colors for this type
    const typeColors = getTypeColors(type);
    const headers = config.getSection('HEADERS');
    
    // Add Type header row with appropriate color
    sheet.getRange(rowIndex, 1).setValue(type);
    sheet.getRange(rowIndex, 1, 1, headers.length)
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
   * @private
   */
  function addCategoryRows(sheet, combinations, rowIndex, type, columnIndices) {
    if (combinations.length === 0) return rowIndex;
    
    const expenseTypes = config.getSection('EXPENSE_TYPES');
    const colors = config.getSection('COLORS').UI;
    const sheetNames = config.getSection('SHEETS');
    
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
    if (expenseTypes.includes(type)) {
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
          sheetName: sheetNames.TRANSACTIONS,
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
      if (config.getSection('PERFORMANCE').USE_BATCH_OPERATIONS) {
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
        
        if (type === config.getSection('TRANSACTION_TYPES').INCOME) {
          cell.setFontColor(colors.INCOME_FONT);
        } else if (expenseTypes.includes(type)) {
          cell.setFontColor(colors.EXPENSE_FONT);
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
   * @private
   */
  function addTypeSubtotalRow(sheet, type, rowIndex, rowCount) {
    // Get colors for this type
    const typeColors = getTypeColors(type);
    const headers = config.getSection('HEADERS');
    
    // Add subtotal for this type
    sheet.getRange(rowIndex, 1).setValue(`Total ${type}`);
    sheet.getRange(rowIndex, 1, 1, headers.length)
      .setBackground(typeColors.BG)
      .setFontWeight("bold")
      .setFontColor(typeColors.FONT);
    
    // Add subtotal formulas for each month column using batch operations
    const formulas = [];
    const startRow = rowIndex - rowCount;
    const endRow = rowIndex - 1;
    
    for (let monthCol = 5; monthCol <= 17; monthCol++) {
      formulas.push(`=SUM(${utils.columnToLetter(monthCol)}${startRow}:${utils.columnToLetter(monthCol)}${endRow})`);
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
   * @private
   */
  function addNetCalculations(sheet, startRow) {
    // Find rows containing total Income, Expenses, and Savings
    const data = sheet.getDataRange().getValues();
    const headers = config.getSection('HEADERS');
    const colors = config.getSection('COLORS').UI;
    
    // Find the rows containing the totals we need
    const totals = findTotalRows(data);
    const { incomeRow, expensesRow, savingsRow } = totals;
    
    if (!incomeRow || !expensesRow) {
      // Couldn't find required rows
      return startRow;
    }
    
    // Add a section header for Net Calculations
    sheet.getRange(startRow, 1).setValue("Net Calculations");
    sheet.getRange(startRow, 1, 1, headers.length)
      .setBackground(colors.NET_BG)
      .setFontWeight("bold")
      .setFontColor(colors.NET_FONT);
    
    startRow++;
    
    // Add Net (Income - Expenses) row
    sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses)");
    sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
    
    // Create all formulas at once
    const netFormulas = [];
    for (let col = 5; col <= 17; col++) {
      netFormulas.push(`=${utils.columnToLetter(col)}${incomeRow}-${utils.columnToLetter(col)}${expensesRow}`);
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
        expSavFormulas.push(`=${utils.columnToLetter(col)}${expensesRow}+${utils.columnToLetter(col)}${savingsRow}`);
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
        totalNetFormulas.push(`=${utils.columnToLetter(col)}${incomeRow}-${utils.columnToLetter(col)}${expensesRow}-${utils.columnToLetter(col)}${savingsRow}`);
      }
      
      // Apply formulas in batch
      sheet.getRange(startRow, 5, 1, 13).setFormulas([totalNetFormulas]);
      formatRangeAsCurrency(sheet.getRange(startRow, 5, 1, 13));
    }
    
    // Add a bottom border to the last row
    sheet.getRange(startRow, 1, 1, headers.length).setBorder(
      null, null, true, null, null, null, 
      colors.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
    
    return startRow + 2; // Add space after net calculations
  }
  
  /**
   * Finds the row numbers for total values in the data
   * @param {Array} data - The sheet data
   * @return {Object} Object containing row indices for totals
   * @private
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
   * Formats the overview sheet for better readability
   * @param {SpreadsheetApp.Sheet} sheet - The overview sheet
   * @private
   */
  function formatOverviewSheet(sheet) {
    const lastRow = sheet.getLastRow();
    const headers = config.getSection('HEADERS');
    const uiConfig = config.getSection('UI').COLUMN_WIDTHS;
    const colors = config.getSection('COLORS').UI;
    
    // Set column widths
    sheet.setColumnWidth(1, uiConfig.TYPE);
    sheet.setColumnWidth(2, uiConfig.CATEGORY);
    sheet.setColumnWidth(3, uiConfig.SUBCATEGORY);
    sheet.setColumnWidth(4, uiConfig.SHARED);
    
    // Set month column widths
    for (let i = 5; i <= 16; i++) {
      sheet.setColumnWidth(i, uiConfig.MONTH);
    }
    
    // Set Total and Average column widths
    sheet.setColumnWidth(17, uiConfig.AVERAGE); // Total column
    sheet.setColumnWidth(18, uiConfig.AVERAGE); // Average column
    
    // Set metrics section column widths
    sheet.setColumnWidth(10, uiConfig.EXPENSE_CATEGORY);
    sheet.setColumnWidth(11, uiConfig.AMOUNT);
    sheet.setColumnWidth(12, uiConfig.RATE);
    sheet.setColumnWidth(13, uiConfig.RATE);
    sheet.setColumnWidth(14, uiConfig.RATE);
    sheet.setColumnWidth(15, uiConfig.AMOUNT);
    
    // Add gridlines and borders to improve readability
    const data = sheet.getDataRange().getValues();
    
    // Add bottom borders to total rows
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].startsWith("Total ")) {
        sheet.getRange(i + 1, 1, 1, headers.length).setBorder(
          null, null, true, null, null, null, 
          colors.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      }
    }
  }
  
  /**
   * Gets the colors for a specific transaction type
   * @param {String} type - The transaction type
   * @return {Object} Object containing background and font colors
   * @private
   */
  function getTypeColors(type) {
    const typeHeaders = config.getSection('COLORS').TYPE_HEADERS;
    const transactionTypes = config.getSection('TRANSACTION_TYPES');
    
    let colors = typeHeaders.DEFAULT;
    
    if (type === transactionTypes.INCOME) {
      colors = typeHeaders.INCOME;
    } else if (type === transactionTypes.ESSENTIALS) {
      colors = typeHeaders.ESSENTIALS;
    } else if (type === "Wants/Pleasure") {
      colors = typeHeaders.WANTS_PLEASURE;
    } else if (type === transactionTypes.EXTRA) {
      colors = typeHeaders.EXTRA;
    } else if (type === transactionTypes.SAVINGS) {
      colors = typeHeaders.SAVINGS;
    }
    
    return colors;
  }
  
  /**
   * Formats a range of cells as currency
   * @param {SpreadsheetApp.Range} range - The range to format
   * @private
   */
  function formatRangeAsCurrency(range) {
    utils.formatAsCurrency(range, 
      config.getSection('LOCALE').CURRENCY_SYMBOL, 
      config.getSection('LOCALE').CURRENCY_LOCALE
    );
  }
  
  /**
   * Gets a user preference from the settings sheet
   * @param {String} key - The preference key
   * @param {any} defaultValue - The default value if not found
   * @return {any} The preference value
   * @private
   */
  function getUserPreference(key, defaultValue) {
    try {
      // This will be replaced by SettingsService in future
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const settingsSheet = ss.getSheetByName(config.getSection('SHEETS').SETTINGS);
      
      if (!settingsSheet) return defaultValue;
      
      const data = settingsSheet.getDataRange().getValues();
      
      // Skip header row
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          return data[i][1];
        }
      }
      
      return defaultValue;
    } catch (error) {
      console.warn(`Failed to get user preference ${key}:`, error);
      return defaultValue;
    }
  }
  
  /**
   * Sets a user preference in the settings sheet
   * @param {String} key - The preference key
   * @param {any} value - The preference value
   * @private
   */
  function setUserPreference(key, value) {
    try {
      // This will be replaced by SettingsService in future
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let settingsSheet = ss.getSheetByName(config.getSection('SHEETS').SETTINGS);
      
      if (!settingsSheet) {
        settingsSheet = ss.insertSheet(config.getSection('SHEETS').SETTINGS);
        settingsSheet.getRange("A1:B1").setValues([["Preference", "Value"]]);
        settingsSheet.getRange("A1:B1").setFontWeight("bold");
        settingsSheet.hideSheet();
      }
      
      const data = settingsSheet.getDataRange().getValues();
      
      // Check if key exists
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          settingsSheet.getRange(i + 1, 2).setValue(value);
          return;
        }
      }
      
      // Key doesn't exist, append it
      settingsSheet.appendRow([key, value]);
    } catch (error) {
      console.warn(`Failed to set user preference ${key}:`, error);
    }
  }
  
  /**
   * Builder class for generating the financial overview
   * @class
   * @private
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
      const sheetNames = config.getSection('SHEETS');
      
      // Get or create the Overview sheet and clear it
      this.overviewSheet = utils.getOrCreateSheet(
        this.spreadsheet, 
        sheetNames.OVERVIEW
      );
      clearSheetContent(this.overviewSheet);
      
      // Get transaction sheet
      this.transactionSheet = this.spreadsheet.getSheetByName(
        sheetNames.TRANSACTIONS
      );
      
      if (!this.transactionSheet) {
        throw errorService.create(
          `Required sheet "${sheetNames.TRANSACTIONS}" not found`,
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
      this.categoryCombinations = cacheService.get(
        config.getSection('CACHE').KEYS.CATEGORY_COMBINATIONS,
        () => getUniqueCategoryCombinations(
          this.transactionData, 
          this.columnIndices.type, 
          this.columnIndices.category, 
          this.columnIndices.subcategory, 
          this.showSubCategories
        )
      );
      
      // Group categories by type
      this.groupedCombinations = cacheService.get(
        config.getSection('CACHE').KEYS.GROUPED_COMBINATIONS,
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
      config.getSection('TYPE_ORDER').forEach(type => {
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
        ...config.get(),
        // Add any additional config needed by FinancialAnalysisService
        TARGET_RATES: {
          ...config.getSection('TARGET_RATES'),
          WANTS_PLEASURE: config.getSection('TARGET_RATES').WANTS // Map WANTS to WANTS_PLEASURE for compatibility
        }
      };
      
      // This will be replaced with FinancialPlanner.FinancialAnalysis.analyze() in future
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
  // PUBLIC API
  // ============================================================================
  
  return {
    /**
     * Creates a financial overview sheet based on transaction data
     * @return {Object} Object containing information about the built overview
     * @public
     */
    create: function() {
      try {
        uiService.showLoadingSpinner("Generating financial overview...");
        cacheService.invalidateAll(); // Clear cache to ensure fresh data
        
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
        
        uiService.hideLoadingSpinner();
        uiService.showSuccessNotification("Financial overview generated successfully!");
        
        return result;
      } catch (error) {
        uiService.hideLoadingSpinner();
        
        // Process the error
        if (error.name === 'FinancialPlannerError') {
          errorService.log(error);
          uiService.showErrorNotification("Error generating overview", error.message);
        } else {
          // Convert to our custom error format and log
          const wrappedError = errorService.create(
            "Failed to generate financial overview", 
            { originalError: error.message, stack: error.stack, severity: "high" }
          );
          errorService.log(wrappedError);
          uiService.showErrorNotification("Error generating overview", error.message);
        }
        
        throw error;
      }
    },
    
    /**
     * Handles edits to the overview sheet, specifically for the sub-category toggle checkbox
     * Must be triggered from the onEdit(e) function
     * @param {Object} e - The edit event object
     * @public
     */
    handleEdit: function(e) {
      // Check if the edit was in the Overview sheet
      if (e.range.getSheet().getName() !== config.getSection('SHEETS').OVERVIEW) return;
      
      // Check if the edit was to the checkbox cell
      const subcategoryToggle = config.getSection('UI').SUBCATEGORY_TOGGLE;
      if (e.range.getA1Notation() === subcategoryToggle.CHECKBOX_CELL) {
        const newValue = e.range.getValue();
        
        // Update the user preference
        setUserPreference("ShowSubCategories", newValue);
        
        // Show loading toast
        uiService.showLoadingSpinner("Updating overview...");
        
        // Regenerate the overview
        try {
          this.create();
          
          const status = newValue ? "showing" : "hiding";
          uiService.showSuccessNotification(`Overview updated, ${status} sub-categories`);
        } catch (error) {
          uiService.showErrorNotification("Update failed", error.message);
        }
      }
    }
  };
})(FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.CacheService, FinancialPlanner.ErrorService, FinancialPlanner.Config);

// ============================================================================
// BACKWARD COMPATIBILITY LAYER
// ============================================================================

/**
 * Creates a financial overview sheet based on transaction data
 * This function is maintained for backward compatibility
 * @return {Object} Object containing information about the built overview
 */
function createFinancialOverview() {
  return FinancialPlanner.FinanceOverview.create();
}

/**
 * Handles edits to the overview sheet, specifically for the sub-category toggle checkbox
 * This function is maintained for backward compatibility
 * Must be triggered from the onEdit(e) function
 * @param {Object} e - The edit event object
 */
function handleOverviewSheetEdits(e) {
  return FinancialPlanner.FinanceOverview.handleEdit(e);
}
