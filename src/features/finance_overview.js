/**
 * Creates a financial overview sheet based on transaction data
 * This function will generate a complete overview sheet with dynamic categories
 * and optional sub-category display based on user preference
 */
function createFinancialOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create the Overview sheet
  const overviewSheet = getOrCreateSheet(ss, "Overview");
  
  // Get transaction and dropdown sheets
  const transactionSheet = ss.getSheetByName("Transactions");
  const dropdownSheet = ss.getSheetByName("Dropdowns");
  
  if (!transactionSheet || !dropdownSheet) {
    SpreadsheetApp.getUi().alert("Error: Could not find required sheets (Transactions or Dropdowns)");
    return;
  }
  
  // Get the user preference for showing sub-categories
  const showSubCategories = getUserPreference("ShowSubCategories", true);
  
  // Get all transaction data
  const transactionData = transactionSheet.getDataRange().getValues();
  const headers = transactionData[0];
  
  // Find required column indices
  const typeIndex = headers.indexOf("Type");
  const categoryIndex = headers.indexOf("Category");
  const subcategoryIndex = headers.indexOf("Sub-Category");
  const dateIndex = headers.indexOf("Date");
  const amountIndex = headers.indexOf("Amount");
  const sharedIndex = headers.indexOf("Shared");
  
  if (typeIndex < 0 || categoryIndex < 0 || subcategoryIndex < 0 || 
      dateIndex < 0 || amountIndex < 0) {
    SpreadsheetApp.getUi().alert("Error: Could not find required columns in Transaction sheet");
    return;
  }
  
  // Set up header row in Overview sheet
  setupHeaderRow(overviewSheet);
  
  // Get unique combinations of Type/Category/Subcategory
  const categoryCombinations = getUniqueCategoryCombinations(transactionData, typeIndex, categoryIndex, subcategoryIndex);
  
  // Add rows for each category combination
  let rowIndex = 2; // Start after header
  let currentType = "";
  
  // Group by Type
  const groupedCombinations = groupCategoryCombinations(categoryCombinations);
  
  // For each type (Income, Expenses, etc.)
  Object.keys(groupedCombinations).forEach(type => {
    // Add Type header row
    overviewSheet.getRange(rowIndex, 1).setValue(type);
    overviewSheet.getRange(rowIndex, 1, 1, 8).setBackground("#f3f3f3").setFontWeight("bold");
    rowIndex++;
    
    // Add rows for each category/subcategory in this type
    groupedCombinations[type].forEach(combo => {
      overviewSheet.getRange(rowIndex, 1).setValue(combo.type);
      overviewSheet.getRange(rowIndex, 2).setValue(combo.category);
      overviewSheet.getRange(rowIndex, 3).setValue(combo.subcategory);
      
      // Add formula for each month column (columns 4-15 for Jan-Dec)
      for (let monthCol = 4; monthCol <= 15; monthCol++) {
        const monthDate = getMonthDateFromColIndex(monthCol);
        const monthFormula = buildMonthlySumFormula(
          combo.type, 
          combo.category, 
          combo.subcategory, 
          monthDate, 
          "Transactions", 
          typeIndex + 1, 
          categoryIndex + 1, 
          subcategoryIndex + 1, 
          dateIndex + 1, 
          amountIndex + 1,
          sharedIndex + 1
        );
        overviewSheet.getRange(rowIndex, monthCol).setFormula(monthFormula);
      }
      
      // Add average formula in column 16 that properly accounts for empty months
      overviewSheet.getRange(rowIndex, 16).setFormula(`=SUM(D${rowIndex}:O${rowIndex})/12`);
      
      rowIndex++;
    });
    
    // Add subtotal for this type
    overviewSheet.getRange(rowIndex, 1).setValue(`Total ${type}`);
    overviewSheet.getRange(rowIndex, 1, 1, 3).setBackground("#e6e6e6").setFontWeight("bold");
    
    // Add subtotal formulas for each month column
    for (let monthCol = 4; monthCol <= 16; monthCol++) {
      const startRow = rowIndex - groupedCombinations[type].length;
      const endRow = rowIndex - 1;
      overviewSheet.getRange(rowIndex, monthCol).setFormula(`=SUM(${columnToLetter(monthCol)}${startRow}:${columnToLetter(monthCol)}${endRow})`);
    }
    
    rowIndex += 2; // Add space between categories
  });
  
  // Add net calculations
  addNetCalculations(overviewSheet, rowIndex);
  
  // Add key metrics section
  addKeyMetricsSection(overviewSheet);
  
  // Format the overview sheet
  formatOverviewSheet(overviewSheet);
  
  // Dynamically show or hide sub-categories column based on user preference
  if (showSubCategories) {
    overviewSheet.showColumns(3, 1); // Show Sub-Category column (column 3, 1 column wide)
  } else {
    overviewSheet.hideColumns(3, 1); // Hide Sub-Category column (column 3, 1 column wide)
  }
}

/**
 * Sets up the header row in the overview sheet and adds the sub-category toggle checkbox
 */
function setupHeaderRow(sheet) {
  const headers = ["Type", "Category", "Sub-Category", "Jan-24", "Feb-24", "Mar-24", "Apr-24", "May-24", "Jun-24", "Jul-24", "Aug-24", "Sep-24", "Oct-24", "Nov-24", "Dec-24", "Average"];
  
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(1, i + 1).setValue(headers[i]);
  }
  
  // Format header row
  sheet.getRange(1, 1, 1, headers.length).setBackground("#d9d9d9").setFontWeight("bold");
  
  // Format the month columns for better readability
  for (let i = 4; i <= 15; i++) {
    sheet.getRange(1, i).setTextRotation(90);
  }
  
  // Add the sub-category toggle checkbox in a more visible location
  const showSubCategories = getUserPreference("ShowSubCategories", true);
  
  // Create a separate section for the checkbox
  const label = sheet.getRange("R1");
  label.setValue("Show Sub-Categories");
  label.setFontWeight("bold");
  
  // Add the checkbox in cell S1 (after the label)
  const checkbox = sheet.getRange("S1");
  checkbox.insertCheckboxes();
  checkbox.setValue(showSubCategories);
  
  // Add a note to explain what the checkbox does
  checkbox.setNote("Toggle to show or hide sub-categories in the overview sheet");
}

/**
 * Handles edits to the overview sheet, specifically for the sub-category toggle checkbox
 * Must be triggered from the onEdit(e) function
 */
function handleOverviewSheetEdits(e) {
  // Check if the edit was in the Overview sheet
  if (e.range.getSheet().getName() !== "Overview") return;
  
  // Check if the edit was to the checkbox cell (S1)
  if (e.range.getA1Notation() === "S1") {
    const newValue = e.range.getValue();
    
    // Update the user preference
    setUserPreference("ShowSubCategories", newValue);
    
    // Regenerate the overview with the new setting
    SpreadsheetApp.getActiveSpreadsheet().toast("Updating overview...", "Please wait");
    
    // Use setTimeout to allow the toast to display before regenerating
    // Note: In Google Apps Script, this is implemented differently
    createFinancialOverview();
    
    const status = newValue ? "showing" : "hiding";
    SpreadsheetApp.getActiveSpreadsheet().toast(`Overview updated, ${status} sub-categories`, "Complete");
  }
}

/**
 * Gets unique combinations of Type/Category/Subcategory from transaction data
 * @param {Array} data - Transaction data
 * @param {Number} typeCol - Column index for transaction type
 * @param {Number} categoryCol - Column index for category
 * @param {Number} subcategoryCol - Column index for subcategory
 * @param {Boolean} showSubCategories - Whether to show subcategories or aggregate by category
 * @return {Array} List of unique category combinations
 */
function getUniqueCategoryCombinations(data, typeCol, categoryCol, subcategoryCol) {
  const combinations = [];
  const seen = new Set();
  const showSubCategories = getUserPreference("ShowSubCategories", true);
  
  // For aggregating by category only (when not showing subcategories)
  const categoryTotals = {};
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = row[typeCol];
    const category = row[categoryCol];
    const subcategory = showSubCategories ? row[subcategoryCol] : "";
    
    if (!type || !category) continue;
    
    const key = `${type}|${category}|${subcategory || ""}`;
    
    if (!seen.has(key)) {
      seen.add(key);
      combinations.push({
        type: type,
        category: category,
        subcategory: subcategory || ""
      });
    }
  }
  
  return combinations;
}

/**
 * Groups category combinations by type
 */
function groupCategoryCombinations(combinations) {
  const grouped = {};
  
  combinations.forEach(combo => {
    if (!grouped[combo.type]) {
      grouped[combo.type] = [];
    }
    grouped[combo.type].push(combo);
  });
  
  // Sort each group by category
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
 * Gets a Date object for the month represented by a column index
 */
function getMonthDateFromColIndex(colIndex) {
  // Column 4 = Jan 2024, 5 = Feb 2024, etc.
  const monthOffset = colIndex - 4;
  return new Date(2024, 0 + monthOffset, 1); // 0 = January (0-indexed)
}

/**
 * Builds a formula to sum transactions for a specific month
 */
function buildMonthlySumFormula(type, category, subcategory, monthDate, sheetName, typeCol, categoryCol, subcategoryCol, dateCol, amountCol, sharedCol) {
  const month = monthDate.getMonth() + 1; // 1-indexed month
  const year = monthDate.getFullYear();
  
  // Calculate the start and end dates for the month
  const startDate = new Date(year, month - 1, 1); // Month is 0-indexed in Date constructor
  const endDate = new Date(year, month, 0); // Last day of the month
  
  // Format dates for Google Sheets (yyyy-mm-dd)
  const startDateFormatted = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const endDateFormatted = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // Create the sum range
  const sumRange = `${sheetName}!${columnToLetter(amountCol)}:${columnToLetter(amountCol)}`;
  
  // Create the criteria pairs for SUMIFS
  let formula = `SUMIFS(${sumRange}`;
  
  // Add type criteria
  formula += `, ${sheetName}!${columnToLetter(typeCol)}:${columnToLetter(typeCol)}, "${type}"`;
  
  // Add category criteria
  formula += `, ${sheetName}!${columnToLetter(categoryCol)}:${columnToLetter(categoryCol)}, "${category}"`;
  
  // Add date range criteria (instead of separate month and year)
  formula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, ">=${startDateFormatted}"`;
  formula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, "<=${endDateFormatted}"`;
  
  // Add subcategory criteria if it exists
  if (subcategory) {
    formula += `, ${sheetName}!${columnToLetter(subcategoryCol)}:${columnToLetter(subcategoryCol)}, "${subcategory}"`;
  }
  
  formula += `)`;
  
  // If this is a shared expense category, add logic to divide by 2 when shared column is TRUE
  if (sharedCol > 0) {
    // Non-shared expenses (Shared = FALSE)
    let nonSharedFormula = `SUMIFS(${sumRange}`;
    nonSharedFormula += `, ${sheetName}!${columnToLetter(typeCol)}:${columnToLetter(typeCol)}, "${type}"`;
    nonSharedFormula += `, ${sheetName}!${columnToLetter(categoryCol)}:${columnToLetter(categoryCol)}, "${category}"`;
    nonSharedFormula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, ">=${startDateFormatted}"`;
    nonSharedFormula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, "<=${endDateFormatted}"`;
    if (subcategory) {
      nonSharedFormula += `, ${sheetName}!${columnToLetter(subcategoryCol)}:${columnToLetter(subcategoryCol)}, "${subcategory}"`;
    }
    nonSharedFormula += `, ${sheetName}!${columnToLetter(sharedCol)}:${columnToLetter(sharedCol)}, FALSE)`;
    
    // Shared expenses (Shared = TRUE, divided by 2)
    let sharedFormula = `SUMIFS(${sumRange}`;
    sharedFormula += `, ${sheetName}!${columnToLetter(typeCol)}:${columnToLetter(typeCol)}, "${type}"`;
    sharedFormula += `, ${sheetName}!${columnToLetter(categoryCol)}:${columnToLetter(categoryCol)}, "${category}"`;
    sharedFormula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, ">=${startDateFormatted}"`;
    sharedFormula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, "<=${endDateFormatted}"`;
    if (subcategory) {
      sharedFormula += `, ${sheetName}!${columnToLetter(subcategoryCol)}:${columnToLetter(subcategoryCol)}, "${subcategory}"`;
    }
    sharedFormula += `, ${sheetName}!${columnToLetter(sharedCol)}:${columnToLetter(sharedCol)}, TRUE)/2`;
    
    return `${nonSharedFormula} + (${sharedFormula})`;
  }
  
  return formula;
}

/**
 * Adds net calculation rows to the overview sheet
 */
function addNetCalculations(sheet, startRow) {
  // Find rows containing total Income and total Expenses
  const data = sheet.getDataRange().getValues();
  let incomeRow = -1;
  let expensesRow = -1;
  let savingsRow = -1;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "Total Income") incomeRow = i + 1;
    if (data[i][0] === "Total Expenses") expensesRow = i + 1;
    if (data[i][0] === "Total Savings") savingsRow = i + 1;
  }
  
  if (incomeRow < 0 || expensesRow < 0) {
    // Couldn't find required rows
    return;
  }
  
  // Add Net (Income - Expenses) row
  sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses)");
  sheet.getRange(startRow, 1, 1, 3).setBackground("#e6e6e6").setFontWeight("bold");
  
  // Add formulas for each month column
  for (let monthCol = 4; monthCol <= 8; monthCol++) {
    sheet.getRange(startRow, monthCol).setFormula(`=${columnToLetter(monthCol)}${incomeRow}-${columnToLetter(monthCol)}${expensesRow}`);
  }
  
  startRow++;
  
  // Add Total Expenses + Savings row if we found a savings row
  if (savingsRow > 0) {
    sheet.getRange(startRow, 1).setValue("Total Expenses + Savings");
    sheet.getRange(startRow, 1, 1, 3).setBackground("#e6e6e6").setFontWeight("bold");
    
    for (let monthCol = 4; monthCol <= 8; monthCol++) {
      sheet.getRange(startRow, monthCol).setFormula(`=${columnToLetter(monthCol)}${expensesRow}+${columnToLetter(monthCol)}${savingsRow}`);
    }
    
    startRow++;
    
    // Add Net (Income - Expenses - Savings) row
    sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses - Savings)");
    sheet.getRange(startRow, 1, 1, 3).setBackground("#e6e6e6").setFontWeight("bold");
    
    for (let monthCol = 4; monthCol <= 8; monthCol++) {
      sheet.getRange(startRow, monthCol).setFormula(`=${columnToLetter(monthCol)}${incomeRow}-${columnToLetter(monthCol)}${expensesRow}-${columnToLetter(monthCol)}${savingsRow}`);
    }
  }
}

/**
 * Adds key metrics section to the overview sheet
 */
function addKeyMetricsSection(sheet) {
  // Find the last row with content
  const lastRow = sheet.getLastRow();
  const metricsStartRow = lastRow + 3;
  
  // Find rows containing total Income, Expenses, Savings
  const data = sheet.getDataRange().getValues();
  let incomeRow = -1;
  let expensesRow = -1;
  let savingsRow = -1;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "Total Income") incomeRow = i + 1;
    if (data[i][0] === "Total Expenses") expensesRow = i + 1;
    if (data[i][0] === "Total Savings") savingsRow = i + 1;
  }
  
  // Add Key Metrics header
  sheet.getRange(metricsStartRow, 10).setValue("Key Metrics");
  sheet.getRange(metricsStartRow, 10, 1, 2).setBackground("#d9d9d9").setFontWeight("bold");
  
  // Add Rate header
  sheet.getRange(metricsStartRow, 12).setValue("Rate");
  sheet.getRange(metricsStartRow, 12).setBackground("#d9d9d9").setFontWeight("bold");
  
  // Add metrics rows
  let currentRow = metricsStartRow + 1;
  
  // Savings Rate
  if (incomeRow > 0 && savingsRow > 0) {
    sheet.getRange(currentRow, 10).setValue("Savings Rate");
    sheet.getRange(currentRow, 12).setFormula(`=H${savingsRow}/H${incomeRow}`);
    currentRow++;
  }
  
  // Expenses/Income
  if (incomeRow > 0 && expensesRow > 0) {
    sheet.getRange(currentRow, 10).setValue("Expenses/Income");
    sheet.getRange(currentRow, 12).setFormula(`=H${expensesRow}/H${incomeRow}`);
    currentRow++;
  }
  
  // Extra
  sheet.getRange(currentRow, 10).setValue("Extra");
  currentRow++;
  
  // Total Expenses
  sheet.getRange(currentRow, 10).setValue("Total Expenses");
  sheet.getRange(currentRow, 12).setFormula(`=IFERROR(H${expensesRow}/H${incomeRow}, "N/A")`);
  
  // Add Expense Categories table
  const expenseStartRow = metricsStartRow + 6;
  
  // Add headers
  sheet.getRange(expenseStartRow, 10).setValue("Expense");
  sheet.getRange(expenseStartRow, 11).setValue("Amount");
  sheet.getRange(expenseStartRow, 12).setValue("Rate");
  sheet.getRange(expenseStartRow, 13).setValue("Target Rate");
  sheet.getRange(expenseStartRow, 14).setValue("% change");
  sheet.getRange(expenseStartRow, 15).setValue("Amount change");
  
  sheet.getRange(expenseStartRow, 10, 1, 6).setBackground("#d9d9d9").setFontWeight("bold");
  
  // Find expense categories (these should be subcategories under "Expenses")
  const expenseCategories = [];
  
  // Log all types found in the data for debugging
  const typesFound = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      typesFound.add(data[i][0]);
    }
  }
  Logger.log("Types found in data: " + Array.from(typesFound).join(", "));
  
  // Look for expense categories based on the user's specific expense types
  const expenseTypes = ["Essentials", "Wants/Pleasure", "Extra"];
  
  for (let i = 1; i < data.length; i++) {
    // Check if this row has a type that's considered an expense
    if (expenseTypes.includes(data[i][0]) && data[i][1]) {
      expenseCategories.push({
        category: data[i][1],
        row: i + 1
      });
      Logger.log("Found expense category: " + data[i][1] + " from type: " + data[i][0] + " at row " + (i + 1));
    }
  }
  
  Logger.log("Found " + expenseCategories.length + " expense categories in total");
  
  // Add rows for each expense category
  currentRow = expenseStartRow + 1;
  expenseCategories.forEach(category => {
    sheet.getRange(currentRow, 10).setValue(category.category);
    sheet.getRange(currentRow, 11).setFormula(`=H${category.row}`);
    sheet.getRange(currentRow, 12).setFormula(`=IFERROR(K${currentRow}/H${incomeRow}, 0)`);
    sheet.getRange(currentRow, 13).setValue(0.2); // Default target rate
    sheet.getRange(currentRow, 14).setFormula(`=IFERROR((L${currentRow}-M${currentRow})/M${currentRow}, 0)`);
    sheet.getRange(currentRow, 15).setFormula(`=IFERROR(K${currentRow}-(H${incomeRow}*M${currentRow}), 0)`);
    currentRow++;
  });
  
  // Add Total Expenses row
  sheet.getRange(currentRow, 10).setValue("Total Expenses");
  sheet.getRange(currentRow, 11).setFormula(`=H${expensesRow}`);
  sheet.getRange(currentRow, 12).setFormula(`=IFERROR(K${currentRow}/H${incomeRow}, 0)`);
  sheet.getRange(currentRow, 13).setValue(1); // Target 100%
  sheet.getRange(currentRow, 14).setFormula(`=IFERROR((L${currentRow}-M${currentRow})/M${currentRow}, 0)`);
  sheet.getRange(currentRow, 15).setFormula(`=IFERROR(K${currentRow}-(H${incomeRow}*M${currentRow}), 0)`);
  
  // Create expenditure breakdown chart only if we have expense categories
  if (expenseCategories.length > 0) {
    Logger.log("Creating expenditure chart with " + expenseCategories.length + " categories");
    createExpenditureChart(sheet, expenseStartRow + 1, currentRow - 1, 10);
  } else {
    Logger.log("Skipping chart creation - no expense categories found");
  }
}

/**
 * Creates a pie chart for expenditure breakdown
 */
function createExpenditureChart(sheet, startRow, endRow, categoryCol) {
  const chartBuilder = sheet.newChart();
  
  // Define chart data range (category name and amount)
  const dataRange = sheet.getRange(startRow, categoryCol, endRow - startRow + 1, 2);
  
  // Create a pie chart
  chartBuilder
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setPosition(startRow, categoryCol + 6, 0, 0)
    .setOption('title', 'Expenditure Breakdown')
    .setOption('pieSliceText', 'percentage')
    .setOption('legend', { position: 'right' })
    .setOption('width', 400)
    .setOption('height', 300);
  
  // Add the chart to the sheet
  sheet.insertChart(chartBuilder.build());
}

/**
 * Formats the overview sheet for better readability
 */
function formatOverviewSheet(sheet) {
  // Format currency columns
  const lastRow = sheet.getLastRow();
  const currencyColumns = [4, 5, 6, 7, 8, 11, 15]; // Columns with monetary values
  
  currencyColumns.forEach(col => {
    formatAsCurrency(sheet.getRange(2, col, lastRow - 1, 1));
  });
  
  // Format percentage columns
  const percentColumns = [12, 13, 14]; // Rate columns
  percentColumns.forEach(col => {
    formatAsPercentage(sheet.getRange(2, col, lastRow - 1, 1));
  });
  
  // Adjust column widths
  sheet.setColumnWidth(1, 150); // Type
  sheet.setColumnWidth(2, 150); // Category
  sheet.setColumnWidth(3, 150); // Sub-Category
  sheet.setColumnWidth(10, 150); // Expense category
  
  // Set alternative row colors for readability
  setAlternatingRowColors(sheet, 2, lastRow);
}
