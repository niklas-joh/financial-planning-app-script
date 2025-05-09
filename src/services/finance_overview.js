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

/**
 * @namespace FinancialPlanner.FinanceOverview
 * @description Service responsible for generating a comprehensive financial overview sheet.
 * It processes transaction data, groups it by type and category (optionally sub-category),
 * calculates monthly totals, and presents a formatted summary.
 * @param {FinancialPlanner.Utils} utils - The utility service.
 * @param {FinancialPlanner.UIService} uiService - The UI service for notifications.
 * @param {FinancialPlanner.CacheService} cacheService - The caching service.
 * @param {FinancialPlanner.ErrorService} errorService - The error handling service.
 * @param {FinancialPlanner.Config} config - The global configuration service.
 * @param {FinancialPlanner.SettingsService} settingsService - The settings management service.
 * @param {FinancialPlanner.FinancialAnalysisService} analysisServiceInstance - An instance of the financial analysis service.
 */
FinancialPlanner.FinanceOverview = (function(utils, uiService, cacheService, errorService, config, settingsService, analysisServiceInstance) { // Added settingsService and analysisServiceInstance
  // ============================================================================
  // PRIVATE IMPLEMENTATION
  // ============================================================================
  
  /**
   * Retrieves and processes raw transaction data from the specified sheet.
   * It identifies column indices for key data points (type, category, date, amount, etc.)
   * and validates the presence of required columns.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The transaction sheet object.
   * @return {{data: Array<Array<any>>, indices: object}} An object containing the raw data as a 2D array
   *                                                      and an `indices` object mapping column names to their 0-based index.
   * @throws {FinancialPlannerError} If required columns are missing in the transaction sheet.
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
   * Extracts unique combinations of Type, Category, and optionally Sub-Category from the transaction data.
   * @param {Array<Array<any>>} data - The raw transaction data (2D array, rows as arrays).
   * @param {number} typeCol - The 0-based column index for the transaction type.
   * @param {number} categoryCol - The 0-based column index for the category.
   * @param {number} subcategoryCol - The 0-based column index for the sub-category.
   * @param {boolean} showSubCategories - If true, sub-categories are included in the combination.
   * @return {Array<{type: string, category: string, subcategory: string}>} An array of objects,
   *         each representing a unique combination. Sub-category will be an empty string if not shown or not present.
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
   * Groups an array of category combinations by their 'type' property.
   * Within each type, combinations are sorted alphabetically by category, then by sub-category.
   * @param {Array<{type: string, category: string, subcategory: string}>} combinations - An array of category combination objects.
   * @return {object<string, Array<{type: string, category: string, subcategory: string}>>} An object where keys are transaction types
   *         and values are arrays of sorted combination objects belonging to that type.
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
   * Constructs a Google Sheets `SUMIFS` formula string to sum transaction amounts for a specific
   * type, category, sub-category (optional), and month. Handles shared expenses by dividing their sum by 2.
   * @param {object} params - Parameters for building the formula.
   * @param {string} params.type - The transaction type.
   * @param {string} params.category - The transaction category.
   * @param {string} [params.subcategory] - The transaction sub-category (optional).
   * @param {Date} params.monthDate - A Date object representing any day in the target month.
   * @param {string} params.sheetName - The name of the transactions sheet (e.g., "Transactions").
   * @param {number} params.typeCol - 1-based column number for 'Type' in the transaction sheet.
   * @param {number} params.categoryCol - 1-based column number for 'Category'.
   * @param {number} params.subcategoryCol - 1-based column number for 'Sub-Category'.
   * @param {number} params.dateCol - 1-based column number for 'Date'.
   * @param {number} params.amountCol - 1-based column number for 'Amount'.
   * @param {number} params.sharedCol - 1-based column number for 'Shared' status (in Transactions sheet, now unused for division).
   * @param {number} overviewSheetCurrentRow - The 1-based row number in the Overview sheet where this formula will be placed.
   * @return {string} The `SUMIFS` formula string.
   * @private
   */
  function buildMonthlySumFormula(params, overviewSheetCurrentRow) { // Added overviewSheetCurrentRow
    const {
      type, category, subcategory, monthDate, sheetName,
      typeCol, categoryCol, subcategoryCol, dateCol, amountCol // sharedCol removed from destructuring
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
    
    // Standard SUMIFS formula based on criteria from Transactions sheet
    const sumifsFormula = `SUMIFS(${sumRange}, ${baseCriteria.join(", ")})`;
    
    // For expense types, divide by 2 if the checkbox in column D of the Overview sheet is TRUE for the current row
    const expenseTypes = config.getSection('EXPENSE_TYPES');
    if (expenseTypes.includes(type)) {
      // The divisor is determined by the state of the checkbox in column D of the current row in the Overview sheet
      const divisorFormula = `IF(D${overviewSheetCurrentRow}=TRUE, 2, 1)`;
      return `(${sumifsFormula}) / ${divisorFormula}`;
    }
    
    // For non-expense types (e.g., Income), return the simple SUMIFS
    return sumifsFormula;
  }
  
  /**
   * Formats a JavaScript Date object into a string based on the date format specified in the application configuration.
   * Uses the script's time zone.
   * @param {Date} date - The Date object to format.
   * @return {string} The formatted date string (e.g., "yyyy-MM-dd").
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
   * Clears all content, formatting, and data validations from the given sheet.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object to clear.
   * @return {void}
   * @private
   */
  function clearSheetContent(sheet) {
    sheet.clear(); // Clear existing content
    sheet.clearFormats(); // Clear existing formats
    // Clear check boxes
    sheet.getRange("A1:Z1000").setDataValidation(null);
  }
  
  /**
   * Sets up the header row of the financial overview sheet.
   * This includes setting header titles, styles, a sub-category toggle checkbox, and freezing the row.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The overview sheet object.
   * @param {boolean} showSubCategories - The current preference for showing sub-categories, used to set the checkbox state.
   * @return {void}
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
   * Adds a styled header row for a specific transaction type (e.g., "Income", "Essentials") to the overview sheet.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The overview sheet object.
   * @param {string} type - The name of the transaction type.
   * @param {number} rowIndex - The 1-based row index where this type header should be inserted.
   * @return {number} The next available row index after inserting the type header.
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
   * Adds rows for each category and sub-category combination under a specific transaction type
   * to the overview sheet. This includes setting their names, shared expense checkboxes (if applicable),
   * and monthly sum formulas.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The overview sheet object.
   * @param {Array<{type: string, category: string, subcategory: string}>} combinations - An array of category combination objects for the current type.
   * @param {number} rowIndex - The 1-based starting row index for inserting these category rows.
   * @param {string} type - The current transaction type being processed.
   * @param {object} columnIndices - An object mapping transaction data column names (e.g., 'type', 'amount') to their 0-based indices.
   * @return {number} The next available row index after inserting all category rows.
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
          sharedCol: columnIndices.shared + 1 // This param is still passed but ignored by new buildMonthlySumFormula division logic
        };
        
        rowFormulas.push(buildMonthlySumFormula(formulaParams, currentRow)); // Pass currentRow
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
    const valueRange = sheet.getRange(startRow, 5, numRows, 13); // Columns E through Q (Total) - Average is col 18 (R)
    
    // Format all values as currency using the default format (which includes [RED])
    formatRangeAsCurrency(valueRange, false); // false indicates not a total row
    
    // Format Average column (R or 18) separately if needed, or include in valueRange if it should also be default currency
    const averageRange = sheet.getRange(startRow, 18, numRows, 1);
    formatRangeAsCurrency(averageRange, false);


    // Color income values (non-total rows)
    // Expense font color is now handled by the number format for negative values
    if (type === config.getSection('TRANSACTION_TYPES').INCOME) {
        for (let i = 0; i < numRows; i++) {
          for (let col = 5; col <= 18; col++) { // E to R
            const cell = sheet.getRange(startRow + i, col);
            cell.setFontColor(colors.INCOME_FONT);
          }
        }
    }
    
    return rowIndex + numRows;
  }
  
  /**
   * Adds a subtotal row for a specific transaction type to the overview sheet.
   * This row sums up the monthly totals for all categories under that type.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The overview sheet object.
   * @param {string} type - The name of the transaction type for which the subtotal is being calculated.
   * @param {number} rowIndex - The 1-based row index where the subtotal row should be inserted.
   * @param {number} rowCount - The number of category/sub-category rows that belong to this type (used for SUM range).
   * @return {number} The next available row index after inserting the subtotal row.
   * @private
   */
  function addTypeSubtotalRow(sheet, type, rowIndex, rowCount) {
    // Get colors for this type
    const typeColors = getTypeColors(type);
    const headers = config.getSection('HEADERS');
    
    // Add subtotal for this type
    sheet.getRange(rowIndex, 1).setValue(`Total ${type}`);
    sheet.getRange(rowIndex, 1, 1, headers.length) // Format the entire row (A to R)
      .setBackground(typeColors.BG)
      .setFontWeight("bold")
      .setFontColor(typeColors.FONT); // Set font color for the whole row first
    
    // Add subtotal formulas for each month column and the average column using batch operations
    const formulas = [];
    const startRowForSum = rowIndex - rowCount;
    const endRowForSum = rowIndex - 1;
    
    // Loop for columns E (5) to R (18)
    for (let monthCol = 5; monthCol <= 18; monthCol++) { 
      formulas.push(`=SUM(${utils.columnToLetter(monthCol)}${startRowForSum}:${utils.columnToLetter(monthCol)}${endRowForSum})`);
    }
    
    // Set all formulas at once for columns E to R (14 columns)
    const formulaRange = sheet.getRange(rowIndex, 5, 1, 14); 
    formulaRange.setFormulas([formulas]);
    
    // Format the subtotal row's numeric cells using CURRENCY_TOTAL_ROW (no [RED])
    formatRangeAsCurrency(formulaRange, true); 
    // Explicitly set font color for numeric cells again to ensure it overrides any number format color
    formulaRange.setFontColor(typeColors.FONT); 
    
    return rowIndex + 1;
  }
  
  /**
   * Adds net calculation rows (e.g., "Net (Total Income - Expenses)") to the overview sheet.
   * It finds the previously calculated total rows for Income, Expenses, and Savings to base these calculations on.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The overview sheet object.
   * @param {number} startRow - The 1-based row index where the net calculations section should begin.
   * @return {number} The next available row index after adding all net calculation rows and a separator.
   * @private
   */
  function addNetCalculations(sheet, startRow) {
    // Find rows containing total Income, Expenses, and Savings
    const data = sheet.getDataRange().getValues();
    const headers = config.getSection('HEADERS');
    const uiColors = config.getSection('COLORS').UI; // Use UI colors
    
    // Find the rows containing the totals we need
    const totals = findTotalRows(data);
    const { incomeRow, expensesRow, savingsRow } = totals;
    
    if (!incomeRow || !expensesRow) {
      // Couldn't find required rows
      return startRow;
    }
    
    // Add a section header for Net Calculations
    sheet.getRange(startRow, 1).setValue("Net Calculations");
    sheet.getRange(startRow, 1, 1, headers.length) // Format entire row
      .setBackground(uiColors.NET_BG)
      .setFontWeight("bold")
      .setFontColor(uiColors.NET_FONT); // Set font for whole row
    
    startRow++;
    
    // Add Net (Income - Expenses) row
    sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses)");
    sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold"); // Label part
    
    const netFormulas = [];
    for (let col = 5; col <= 18; col++) { // E to R (Average)
      netFormulas.push(`=${utils.columnToLetter(col)}${incomeRow}-${utils.columnToLetter(col)}${expensesRow}`);
    }
    
    const netNumericRange = sheet.getRange(startRow, 5, 1, 14); // E to R
    netNumericRange.setFormulas([netFormulas]);
    formatRangeAsCurrency(netNumericRange, true); // Use total row format
    netNumericRange.setFontColor(uiColors.NET_FONT); // Ensure font color for numbers
    
    startRow++;
    
    if (savingsRow) {
      // Add Total Expenses + Savings row
      sheet.getRange(startRow, 1).setValue("Total Expenses + Savings");
      sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
      
      const expSavFormulas = [];
      for (let col = 5; col <= 18; col++) { // E to R
        expSavFormulas.push(`=${utils.columnToLetter(col)}${expensesRow}+${utils.columnToLetter(col)}${savingsRow}`);
      }
      
      const expSavNumericRange = sheet.getRange(startRow, 5, 1, 14); // E to R
      expSavNumericRange.setFormulas([expSavFormulas]);
      formatRangeAsCurrency(expSavNumericRange, true);
      expSavNumericRange.setFontColor(uiColors.NET_FONT);
      
      startRow++;
      
      // Add Net (Income - Expenses - Savings) row
      sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses - Savings)");
      sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
      
      const totalNetFormulas = [];
      for (let col = 5; col <= 18; col++) { // E to R
        totalNetFormulas.push(`=${utils.columnToLetter(col)}${incomeRow}-${utils.columnToLetter(col)}${expensesRow}-${utils.columnToLetter(col)}${savingsRow}`);
      }
      
      const totalNetNumericRange = sheet.getRange(startRow, 5, 1, 14); // E to R
      totalNetNumericRange.setFormulas([totalNetFormulas]);
      formatRangeAsCurrency(totalNetNumericRange, true);
      totalNetNumericRange.setFontColor(uiColors.NET_FONT);
    }
    
    // Add a bottom border to the last row
    sheet.getRange(startRow, 1, 1, headers.length).setBorder(
      null, null, true, null, null, null, 
      uiColors.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
    
    return startRow + 2; // Add space after net calculations
  }
  
  /**
   * Scans the sheet data to find the 1-based row numbers for "Total Income", "Total Expenses", and "Total Savings".
   * @param {Array<Array<any>>} data - A 2D array representing the data from the overview sheet.
   * @return {{incomeRow: number|null, expensesRow: number|null, savingsRow: number|null}} An object containing
   *         the 1-based row indices. Values are null if a corresponding total row is not found.
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
      if (data[i][0] === "Total Expenses") totals.expensesRow = i + 1; // This will capture the "Total Expenses" specific row if it exists
      if (data[i][0] === "Total Savings") totals.savingsRow = i + 1;
    }
    
    // If "Total Expenses" wasn't found directly, try to find the last expense type total row
    // This logic might need adjustment if "Total Expenses" is explicitly added elsewhere
    if (!totals.expensesRow) {
        const expenseTypes = config.getSection('EXPENSE_TYPES');
        let lastExpenseTypeRow = -1;
        for (let i = data.length - 1; i >= 0; i--) {
            if (expenseTypes.some(et => data[i][0] === `Total ${et}`)) {
                lastExpenseTypeRow = i + 1;
                break;
            }
        }
        if (lastExpenseTypeRow !== -1) {
            // This is a fallback, ideally "Total Expenses" row is explicitly created
            // For now, let's assume the net calculations need a row that sums all expenses.
            // The current structure sums individual expense type totals.
            // If a single "Total Expenses" row is needed for net calculations, it should be created.
            // For now, we'll use the last found expense type total as a proxy if "Total Expenses" is missing.
            // This part of the logic might be complex if "Total Expenses" isn't a single, clearly identifiable row.
        }
    }


    return totals;
  }
  
  /**
   * Applies various formatting options to the overview sheet for better readability,
   * including setting column widths and adding borders to total rows.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The overview sheet object.
   * @return {void}
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
    
    // Set month column widths (E to P, columns 5 to 16)
    for (let i = 5; i <= 16; i++) {
      sheet.setColumnWidth(i, uiConfig.MONTH);
    }
    
    // Set Total and Average column widths (Q and R, columns 17 and 18)
    sheet.setColumnWidth(17, uiConfig.AVERAGE); // Total column
    sheet.setColumnWidth(18, uiConfig.AVERAGE); // Average column
    
    // Add gridlines and borders to improve readability
    const data = sheet.getDataRange().getValues();
    
    // Add bottom borders to total rows
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().startsWith("Total ")) {
        sheet.getRange(i + 1, 1, 1, headers.length).setBorder( // headers.length should be 18
          null, null, true, null, null, null, 
          colors.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      }
    }
  }
  
  /**
   * Retrieves the configured background and font colors for a given transaction type.
   * Uses default colors if a specific type is not found in the configuration.
   * @param {string} type - The name of the transaction type.
   * @return {{BG: string, FONT: string}} An object containing `BG` (background color hex) and `FONT` (font color hex).
   * @private
   */
  function getTypeColors(type) {
    const typeHeaders = config.getSection('COLORS').TYPE_HEADERS;
    const transactionTypes = config.getSection('TRANSACTION_TYPES');
    
    let colors = typeHeaders.DEFAULT; // Default
    
    // Find a match in a case-insensitive way, or use direct mapping
    const normalizedType = type.toLowerCase();
    for (const key in transactionTypes) {
        if (transactionTypes[key].toLowerCase() === normalizedType && typeHeaders[key]) {
            colors = typeHeaders[key];
            break;
        }
    }
    // Fallback for "Wants/Pleasure" if not directly mapped via ENUM key
    if (normalizedType === "wants/pleasure" && typeHeaders.WANTS_PLEASURE) {
        colors = typeHeaders.WANTS_PLEASURE;
    }
    
    return colors;
  }
  
  /**
   * Formats a given cell range as currency using the appropriate number format string from the application configuration.
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell range to format.
   * @param {boolean} [isTotalRow=false] - Whether this range is for a total row, which uses a different number format.
   * @return {void}
   * @private
   */
  function formatRangeAsCurrency(range, isTotalRow = false) {
    const localeConfig = config.getSection('LOCALE');
    const numberFormatString = isTotalRow ? 
      localeConfig.NUMBER_FORMATS.CURRENCY_TOTAL_ROW : 
      localeConfig.NUMBER_FORMATS.CURRENCY_DEFAULT;
    utils.formatAsCurrency(range, numberFormatString); // utils.formatAsCurrency now expects the full format string
  }
  
  /**
   * Retrieves a user preference value using the `SettingsService`.
   * Logs an error and returns `defaultValue` if retrieval fails.
   * @param {string} key - The preference key.
   * @param {any} defaultValue - The default value to return if the preference is not found or an error occurs.
   * @return {any} The preference value or `defaultValue`.
   * @private
   */
  function getUserPreference(key, defaultValue) {
    try {
      return settingsService.getValue(key, defaultValue);
    } catch (error) {
      if (errorService && errorService.log) {
        errorService.log(errorService.create(`Failed to get user preference '${key}'`, { originalError: error.toString(), severity: "medium" }));
      } else {
        console.warn(`Failed to get user preference '${key}':`, error.toString());
      }
      return defaultValue;
    }
  }
  
  /**
   * Sets a user preference value using the `SettingsService`.
   * Logs an error if setting the preference fails.
   * @param {string} key - The preference key.
   * @param {any} value - The value to set for the preference.
   * @return {void}
   * @private
   */
  function setUserPreference(key, value) {
    try {
      settingsService.setValue(key, value);
    } catch (error) {
      if (errorService && errorService.log) {
        errorService.log(errorService.create(`Failed to set user preference '${key}'`, { originalError: error.toString(), valueToSet: value, severity: "medium" }));
      } else {
        console.warn(`Failed to set user preference '${key}':`, error.toString());
      }
    }
  }
  
  /**
   * Implements the Builder pattern for constructing the financial overview sheet step-by-step.
   * Encapsulates the state and logic required for overview generation.
   * @class FinancialOverviewBuilder
   * @private
   */
  class FinancialOverviewBuilder {
    /**
     * Initializes the builder's state variables.
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
     * Initializes the builder by getting references to the active spreadsheet,
     * the overview sheet (creating and clearing it if necessary), and the transaction sheet.
     * Also retrieves the user preference for showing sub-categories.
     * @return {FinancialOverviewBuilder} Returns the builder instance for method chaining.
     * @throws {FinancialPlannerError} If the required 'Transactions' sheet is not found.
     */
    initialize() {
      this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheetNames = config.getSection('SHEETS');
      
      this.overviewSheet = utils.getOrCreateSheet(
        this.spreadsheet, 
        sheetNames.OVERVIEW
      );
      clearSheetContent(this.overviewSheet);
      
      this.transactionSheet = this.spreadsheet.getSheetByName(
        sheetNames.TRANSACTIONS
      );
      
      if (!this.transactionSheet) {
        throw errorService.create(
          `Required sheet "${sheetNames.TRANSACTIONS}" not found`,
          { severity: "high" }
        );
      }
      
      this.showSubCategories = getUserPreference("ShowSubCategories", true);
      
      return this;
    }
    
    /**
     * Processes the transaction data by retrieving it, identifying column indices,
     * extracting unique category combinations (using cache), and grouping them by type (using cache).
     * Stores the processed data and combinations within the builder instance.
     * @return {FinancialOverviewBuilder} Returns the builder instance for method chaining.
     */
    processData() {
      const { data, indices } = getProcessedTransactionData(this.transactionSheet);
      this.transactionData = data;
      this.columnIndices = indices;
      
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
      
      this.groupedCombinations = cacheService.get(
        config.getSection('CACHE').KEYS.GROUPED_COMBINATIONS,
        () => groupCategoryCombinations(this.categoryCombinations)
      );
      
      return this;
    }
    
    /**
     * Sets up the header row on the overview sheet using the `setupHeaderRow` helper function.
     * @return {FinancialOverviewBuilder} Returns the builder instance for method chaining.
     */
    setupHeader() {
      setupHeaderRow(this.overviewSheet, this.showSubCategories);
      return this;
    }
    
    /**
     * Generates the main body of the overview sheet.
     * Iterates through transaction types in the configured order, adding type headers,
     * category/sub-category rows with formulas, and type subtotals.
     * @return {FinancialOverviewBuilder} Returns the builder instance for method chaining.
     */
    generateContent() {
      let rowIndex = 2; 
      
      config.getSection('TYPE_ORDER').forEach(type => {
        if (!this.groupedCombinations[type]) return;
        
        rowIndex = addTypeHeaderRow(this.overviewSheet, type, rowIndex);
        
        rowIndex = addCategoryRows(
          this.overviewSheet, 
          this.groupedCombinations[type], 
          rowIndex, 
          type, 
          this.columnIndices
        );
        
        rowIndex = addTypeSubtotalRow(
          this.overviewSheet, 
          type, 
          rowIndex, 
          this.groupedCombinations[type].length
        );
        
        rowIndex += 2; 
      });
      
      this.lastContentRowIndex = rowIndex;
      return this;
    }
    
    /**
     * Adds the 'Net Calculations' section to the overview sheet using the `addNetCalculations` helper function.
     * @return {FinancialOverviewBuilder} Returns the builder instance for method chaining.
     */
    addNetCalculations() {
      this.lastContentRowIndex = addNetCalculations(
        this.overviewSheet, 
        this.lastContentRowIndex
      );
      return this;
    }
    
    /**
     * Adds the financial metrics section to the overview sheet by calling the `analyze` method
     * of the injected `FinancialAnalysisService`. Logs an error if the service is unavailable.
     * @return {FinancialOverviewBuilder} Returns the builder instance for method chaining.
     */
    addMetrics() {
      if (analysisServiceInstance && analysisServiceInstance.analyze) {
         analysisServiceInstance.analyze(this.spreadsheet, this.overviewSheet);
      } else {
        console.error("FinancialAnalysisService not available for addMetrics");
        if (errorService) {
            errorService.log(errorService.create("FinancialAnalysisService not available in FinanceOverview", { severity: "high"}));
        }
      }
      
      return this;
    }
    
    /**
     * Applies formatting (column widths, borders) to the overview sheet using the `formatOverviewSheet` helper function.
     * @return {FinancialOverviewBuilder} Returns the builder instance for method chaining.
     */
    formatSheet() {
      formatOverviewSheet(this.overviewSheet);
      return this;
    }
    
    /**
     * Applies user preferences to the sheet, specifically showing or hiding the sub-category column
     * based on the `showSubCategories` state determined during initialization.
     * @return {FinancialOverviewBuilder} Returns the builder instance for method chaining.
     */
    applyPreferences() {
      if (this.showSubCategories) {
        this.overviewSheet.showColumns(3, 1);
      } else {
        this.overviewSheet.hideColumns(3, 1);
      }
      return this;
    }
    
    /**
     * Finalizes the build process and returns information about the generated overview.
     * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number, success: boolean}} An object containing
     *         a reference to the overview sheet, the last row number with content, and a success flag.
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
     * Creates or regenerates the financial overview sheet.
     * This is the main public entry point for generating the overview.
     * It utilizes the `FinancialOverviewBuilder` to perform the steps.
     * Provides UI feedback (loading spinner, success/error messages).
     * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number, success: boolean}} An object containing
     *         a reference to the overview sheet, the last row number with content, and a success flag.
     * @throws {Error} Re-throws any error encountered during the build process after logging and notifying the user.
     * @public
     * @example
     * // Called from a menu item or script:
     * FinancialPlanner.FinanceOverview.create();
     */
    create: function() {
      try {
        uiService.showLoadingSpinner("Generating financial overview...");
        cacheService.invalidateAll(); 
        
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
        
        if (error.name === 'FinancialPlannerError') {
          errorService.log(error);
          uiService.showErrorNotification("Error generating overview", error.message);
        } else {
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
     * Handles edit events specifically for the 'Overview' sheet.
     * Currently, it only reacts to changes in the 'Show Sub-Categories' checkbox.
     * If the checkbox state changes, it updates the user preference and regenerates the overview.
     * Intended to be called from a central `onEdit` dispatcher (like `FinancialPlanner.Controllers.onEdit`).
     * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The edit event object provided by Google Apps Script.
     *        See {@link https://developers.google.com/apps-script/guides/triggers/events#edit_3}
     * @return {void}
     * @public
     */
    handleEdit: function(e) {
      try {
        if (e.range.getSheet().getName() !== config.getSection('SHEETS').OVERVIEW) return;

        const subcategoryToggle = config.getSection('UI').SUBCATEGORY_TOGGLE;
        if (e.range.getA1Notation() === subcategoryToggle.CHECKBOX_CELL) {
          const newValue = e.range.getValue(); 

          setUserPreference("ShowSubCategories", newValue);
          uiService.showLoadingSpinner("Updating overview based on preference change...");
          this.create(); 
        }
      } catch (error) {
         errorService.handle(errorService.create("Error handling Overview sheet edit", { originalError: error.toString(), eventDetails: JSON.stringify(e) }), "Failed to process change on Overview sheet.");
      }
    }
  };
})(
  FinancialPlanner.Utils, 
  FinancialPlanner.UIService, 
  FinancialPlanner.CacheService, 
  FinancialPlanner.ErrorService, 
  FinancialPlanner.Config,
  FinancialPlanner.SettingsService, 
  FinancialPlanner.FinancialAnalysisService 
);

// ============================================================================
// BACKWARD COMPATIBILITY LAYER
// ============================================================================

/**
 * Creates the financial overview sheet.
 * Maintained for backward compatibility with older triggers or direct calls.
 * Delegates to `FinancialPlanner.FinanceOverview.create()`.
 * @return {object | undefined} The result object from `FinancialPlanner.FinanceOverview.create()`, or undefined if the service isn't loaded.
 * @global
 */
function createFinancialOverview() {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.FinanceOverview && FinancialPlanner.FinanceOverview.create) {
    return FinancialPlanner.FinanceOverview.create();
  }
  Logger.log("Global createFinancialOverview: FinancialPlanner.FinanceOverview not available.");
}

/**
 * Handles edits on the overview sheet.
 * Maintained for backward compatibility with older `onEdit` triggers that might call this directly.
 * Delegates to `FinancialPlanner.FinanceOverview.handleEdit(e)`.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The edit event object.
 * @global
 * @return {void}
 */
function handleOverviewSheetEdits(e) {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.FinanceOverview && FinancialPlanner.FinanceOverview.handleEdit) {
    FinancialPlanner.FinanceOverview.handleEdit(e);
  } else {
    Logger.log("Global handleOverviewSheetEdits: FinancialPlanner.FinanceOverview not available.");
  }
}
