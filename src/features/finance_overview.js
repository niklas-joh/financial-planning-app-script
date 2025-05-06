/**
 * Creates a financial overview sheet based on transaction data
 * This function will generate a complete overview sheet with dynamic categories
 * and optional sub-category display based on user preference
 */
function createFinancialOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create the Overview sheet
  const overviewSheet = getOrCreateSheet(ss, "Overview");
  overviewSheet.clear(); // Clear existing content
  overviewSheet.clearFormats(); // Clear existing formats
  // clear check boxes
  overviewSheet.getRange("A1:Z1000").setDataValidation(null);
  
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
  
  // Define the order of main types
  const typeOrder = ["Income", "Essentials", "Wants/Pleasure", "Extra", "Savings"];
  
  // For each type in the defined order
  typeOrder.forEach(type => {
    // Skip if this type doesn't exist in the data
    if (!groupedCombinations[type]) return;
    // Define colors based on type
    let typeBgColor;
    let typeFontColor = "#FFFFFF"; // White text for all type headers
    
    if (type === "Income") {
      typeBgColor = "#2E7D32"; // Green for Income
    } else if (type === "Essentials") {
      typeBgColor = "#1976D2"; // Blue for Essentials
    } else if (type === "Wants/Pleasure") {
      typeBgColor = "#FFA000"; // Amber for Wants/Pleasure
    } else if (type === "Extra") {
      typeBgColor = "#7B1FA2"; // Purple for Extra
    } else if (type === "Savings") {
      typeBgColor = "#1565C0"; // Blue for Savings
    } else {
      typeBgColor = "#424242"; // Dark gray for other types
    }
    
    // Add Type header row with appropriate color
    overviewSheet.getRange(rowIndex, 1).setValue(type);
    overviewSheet.getRange(rowIndex, 1, 1, 17) // Adjusted for new column count
      .setBackground(typeBgColor)
      .setFontWeight("bold")
      .setFontColor(typeFontColor);
    rowIndex++;
    
    // Define category background colors based on type
    let categoryBgColor;
    let categoryLightBgColor;
    
    if (type === "Income") {
      categoryBgColor = "#C8E6C9"; // Light green for Income categories
      categoryLightBgColor = "#E8F5E9"; // Very light green for subcategories
    } else if (type === "Essentials") {
      categoryBgColor = "#FFCDD2"; // Light red for Expense categories
      categoryLightBgColor = "#FFEBEE"; // Very light red for subcategories
    } else if (type === "Wants/Pleasure") {
      categoryBgColor = "#FFE0B2"; // Light orange for Wants/Pleasure categories
      categoryLightBgColor = "#FFF3E0"; // Very light orange for subcategories
    } else if (type === "Extra") {
      categoryBgColor = "#BBDEFB"; // Light blue for Extra categories
      categoryLightBgColor = "#E3F2FD"; // Very light blue for subcategories
    } else if (type === "Savings") {
      categoryBgColor = "#BBDEFB"; // Light blue for Savings categories
      categoryLightBgColor = "#E3F2FD"; // Very light blue for subcategories
    } else {
      categoryBgColor = "#F5F5F5"; // Light gray for other categories
      categoryLightBgColor = "#FAFAFA"; // Very light gray for subcategories
    }
    
    // Store all expense types in an array
    const expenses = ["Essentials", "Wants/Pleasure", "Extra"];

    
    // Add rows for each category/subcategory in this type
    groupedCombinations[type].forEach(combo => {
      // Set values for Type, Category, Sub-Category
      overviewSheet.getRange(rowIndex, 1).setValue(combo.type);
      overviewSheet.getRange(rowIndex, 2).setValue(combo.category);
      overviewSheet.getRange(rowIndex, 3).setValue(combo.subcategory);
      
      // Set Shared? column value (checkbox)
      if (expenses.includes(combo.type)) {
        // Only show checkbox for these types
        overviewSheet.getRange(rowIndex, 4).insertCheckboxes();
        // We'll leave it unchecked by default, but this could be enhanced to show actual shared status
      }
      
      // Apply styling to the row - no background color for regular rows
      if (combo.subcategory) {
        // This is a subcategory row - no background, just indent
        overviewSheet.getRange(rowIndex, 3).setIndent(5); // Indent subcategory for visual hierarchy
      } else {
        // This is a main category row - no background, just bold category name
        overviewSheet.getRange(rowIndex, 2).setFontWeight("bold"); // Bold category name
      }
      
      // Add formula for each month column (columns 5-16 for Jan-Dec, adjusted for Shared? column)
      for (let monthCol = 5; monthCol <= 16; monthCol++) {
        const monthDate = getMonthDateFromColIndex(monthCol - 1); // Adjust for Shared? column
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
        
        // Apply conditional formatting for positive/negative values
        if (combo.type === "Income") {
          // Income should be displayed in green
          overviewSheet.getRange(rowIndex, monthCol).setFontColor("#388E3C"); // Green for income
        } else if (expenses.includes(combo.type)) {
          // Expenses should be displayed in red
          overviewSheet.getRange(rowIndex, monthCol).setFontColor("#D32F2F"); // Red for expenses
        }
      }
      
      // Add average formula in column 17 that properly accounts for empty months (adjusted for Shared? column)
      overviewSheet.getRange(rowIndex, 17).setFormula(`=AVERAGE(E${rowIndex}:P${rowIndex})`);
      
      rowIndex++;
    });
    
    // Add subtotal for this type
    overviewSheet.getRange(rowIndex, 1).setValue(`Total ${type}`);
    overviewSheet.getRange(rowIndex, 1, 1, 17).setBackground(typeBgColor).setFontWeight("bold").setFontColor(typeFontColor);
    
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
  
  // Format the overview sheet - with modified row coloring
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
  // Define color constants for better consistency
  const HEADER_BG_COLOR = "#C62828"; // Deep red for headers
  const HEADER_TEXT_COLOR = "#FFFFFF"; // White text for better contrast on red
  
  // Updated headers array with Shared? column
  const headers = ["Type", "Category", "Sub-Category", "Shared?", "Jan-24", "Feb-24", "Mar-24", "Apr-24", "May-24", "Jun-24", "Jul-24", "Aug-24", "Sep-24", "Oct-24", "Nov-24", "Dec-24", "Average"];
  
  // Set header values
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(1, i + 1).setValue(headers[i]);
  }
  
  // Format header row with bold red background and white text
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(HEADER_BG_COLOR)
             .setFontWeight("bold")
             .setFontColor(HEADER_TEXT_COLOR)
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
  
  // Add the sub-category toggle checkbox in a more visible location
  const showSubCategories = getUserPreference("ShowSubCategories", true);
  
  // Create a separate section for the checkbox
  const label = sheet.getRange("S1"); // Moved one column to the right due to Shared? column
  label.setValue("Show Sub-Categories");
  label.setFontWeight("bold");
  
  // Add the checkbox in cell T1 (after the label)
  const checkbox = sheet.getRange("T1"); // Moved one column to the right
  checkbox.insertCheckboxes();
  checkbox.setValue(showSubCategories);
  
  // Add a note to explain what the checkbox does
  checkbox.setNote("Toggle to show or hide sub-categories in the overview sheet");
  
  // Freeze the header row so it remains visible when scrolling
  sheet.setFrozenRows(1);
  
  // Set column width for the Shared? column
  sheet.setColumnWidth(4, 80); // Shared? column
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
  
  // Define the order for expense categories
  const expenseCategoryOrder = ["Essentials", "Wants/Pleasure", "Extra"];
  
  // Sort each group by category
  Object.keys(grouped).forEach(type => {
    if (type === "Expenses") {
      // For Expenses, sort by the predefined order
      grouped[type].sort((a, b) => {
        // Get the index of each category in the order array
        const indexA = expenseCategoryOrder.indexOf(a.category);
        const indexB = expenseCategoryOrder.indexOf(b.category);
        
        // If both categories are in the order array, sort by their order
        if (indexA >= 0 && indexB >= 0) {
          return indexA - indexB;
        }
        
        // If only one is in the order array, prioritize it
        if (indexA >= 0) return -1;
        if (indexB >= 0) return 1;
        
        // If neither is in the order array, sort alphabetically
        const categoryCompare = a.category.localeCompare(b.category);
        // Secondary sort by subcategory if categories are the same
        return categoryCompare !== 0 ? categoryCompare : 
               (a.subcategory || "").localeCompare(b.subcategory || "");
      });
    } else {
      // For other types, sort alphabetically
      grouped[type].sort((a, b) => {
        // Primary sort by category
        const categoryCompare = a.category.localeCompare(b.category);
        // Secondary sort by subcategory if categories are the same
        return categoryCompare !== 0 ? categoryCompare : 
               (a.subcategory || "").localeCompare(b.subcategory || "");
      });
    }
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
    // Non-shared expenses (Shared = "")
    let nonSharedFormula = `SUMIFS(${sumRange}`;
    nonSharedFormula += `, ${sheetName}!${columnToLetter(typeCol)}:${columnToLetter(typeCol)}, "${type}"`;
    nonSharedFormula += `, ${sheetName}!${columnToLetter(categoryCol)}:${columnToLetter(categoryCol)}, "${category}"`;
    nonSharedFormula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, ">=${startDateFormatted}"`;
    nonSharedFormula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, "<=${endDateFormatted}"`;
    if (subcategory) {
      nonSharedFormula += `, ${sheetName}!${columnToLetter(subcategoryCol)}:${columnToLetter(subcategoryCol)}, "${subcategory}"`;
    }
    nonSharedFormula += `, ${sheetName}!${columnToLetter(sharedCol)}:${columnToLetter(sharedCol)}, "")`;
    
    // Shared expenses (Shared = TRUE, divided by 2)
    let sharedFormula = `SUMIFS(${sumRange}`;
    sharedFormula += `, ${sheetName}!${columnToLetter(typeCol)}:${columnToLetter(typeCol)}, "${type}"`;
    sharedFormula += `, ${sheetName}!${columnToLetter(categoryCol)}:${columnToLetter(categoryCol)}, "${category}"`;
    sharedFormula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, ">=${startDateFormatted}"`;
    sharedFormula += `, ${sheetName}!${columnToLetter(dateCol)}:${columnToLetter(dateCol)}, "<=${endDateFormatted}"`;
    if (subcategory) {
      sharedFormula += `, ${sheetName}!${columnToLetter(subcategoryCol)}:${columnToLetter(subcategoryCol)}, "${subcategory}"`;
    }
    sharedFormula += `, ${sheetName}!${columnToLetter(sharedCol)}:${columnToLetter(sharedCol)}, "true")/2`;
    
    return `${nonSharedFormula} + (${sharedFormula})`;
  }
  
  return formula;
}

/**
 * Adds net calculation rows to the overview sheet
 */
function addNetCalculations(sheet, startRow) {
  // Define color constants for better consistency
  const NET_BG_COLOR = "#424242"; // Dark gray for net calculations
  const NET_TEXT_COLOR = "#FFFFFF"; // White text for better contrast
  const BORDER_COLOR = "#FF8F00"; // Amber for borders
  
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
  
  // Add a section header for Net Calculations
  sheet.getRange(startRow, 1).setValue("Net Calculations");
  sheet.getRange(startRow, 1, 1, 17)
    .setBackground(NET_BG_COLOR)
    .setFontWeight("bold")
    .setFontColor(NET_TEXT_COLOR);
  
  startRow++;
  
  // Add Net (Income - Expenses) row
  sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses)");
  sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
  
  // Add formulas for each month column (adjusted for Shared? column)
  for (let monthCol = 5; monthCol <= 17; monthCol++) {
    sheet.getRange(startRow, monthCol).setFormula(`=${columnToLetter(monthCol)}${incomeRow}-${columnToLetter(monthCol)}${expensesRow}`);
    
    // Apply conditional formatting (green for positive, red for negative)
    formatAsCurrency(sheet.getRange(startRow, monthCol));
  }
  
  startRow++;
  
  // Add Total Expenses + Savings row if we found a savings row
  if (savingsRow > 0) {
    sheet.getRange(startRow, 1).setValue("Total Expenses + Savings");
    sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
    
    for (let monthCol = 5; monthCol <= 17; monthCol++) {
      sheet.getRange(startRow, monthCol).setFormula(`=${columnToLetter(monthCol)}${expensesRow}+${columnToLetter(monthCol)}${savingsRow}`);
      
      // Apply conditional formatting
      formatAsCurrency(sheet.getRange(startRow, monthCol));
    }
    
    startRow++;
    
    // Add Net (Income - Expenses - Savings) row
    sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses - Savings)");
    sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
    
    for (let monthCol = 5; monthCol <= 17; monthCol++) {
      sheet.getRange(startRow, monthCol).setFormula(`=${columnToLetter(monthCol)}${incomeRow}-${columnToLetter(monthCol)}${expensesRow}-${columnToLetter(monthCol)}${savingsRow}`);
      
      // Apply conditional formatting
      formatAsCurrency(sheet.getRange(startRow, monthCol));
    }
    
    // Add a bottom border to the last net calculation row
    sheet.getRange(startRow, 1, 1, 17).setBorder(
      null, null, true, null, null, null, 
      BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  } else {
    // If no savings row, add a bottom border to the Net (Income - Expenses) row
    sheet.getRange(startRow - 1, 1, 1, 17).setBorder(
      null, null, true, null, null, null, 
      BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  }
}

/**
 * Adds key metrics section to the overview sheet
 */
function addKeyMetricsSection(sheet) {
  // Define color constants for better consistency
  const HEADER_BG_COLOR = "#C62828"; // Deep red for headers
  const HEADER_TEXT_COLOR = "#FFFFFF"; // White text for better contrast
  const METRICS_BG_COLOR = "#FFEBEE"; // Very light red for metrics section
  const BORDER_COLOR = "#FF8F00"; // Amber for borders
  
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
  sheet.getRange(metricsStartRow, 10, 1, 3)
    .setBackground(HEADER_BG_COLOR)
    .setFontWeight("bold")
    .setFontColor(HEADER_TEXT_COLOR)
    .setHorizontalAlignment("center");
  
  // Add metrics rows
  let currentRow = metricsStartRow + 1;
  
  // Create a metrics table with better formatting
  const metricsTable = [
    ["Metric", "Value", "Target"],
  ];
  
  // Add table headers
  sheet.getRange(currentRow, 10, 1, 3)
    .setValues([["Metric", "Value", "Target"]])
    .setBackground("#F5F5F5")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  currentRow++;
  
  // Savings Rate
  if (incomeRow > 0 && savingsRow > 0) {
    sheet.getRange(currentRow, 10).setValue("Savings Rate");
    sheet.getRange(currentRow, 11).setFormula(`=Q${savingsRow}/Q${incomeRow}`); // Using Average column (Q)
    sheet.getRange(currentRow, 12).setValue(0.2); // 20% target
    sheet.getRange(currentRow, 10, 1, 3).setBackground(METRICS_BG_COLOR);
    
    // Format as percentage
    sheet.getRange(currentRow, 11, 1, 2).setNumberFormat("0.0%");
    
    // Add conditional formatting (green if meeting target, red if not)
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(sheet.getRange(currentRow, 12).getValue())
      .setBackground("#FFCDD2") // Light red if below target
      .setRanges([sheet.getRange(currentRow, 11)])
      .build();
    
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
    
    currentRow++;
  }
  
  // Expenses/Income Ratio
  if (incomeRow > 0 && expensesRow > 0) {
    sheet.getRange(currentRow, 10).setValue("Expenses/Income Ratio");
    sheet.getRange(currentRow, 11).setFormula(`=Q${expensesRow}/Q${incomeRow}`); // Using Average column (Q)
    sheet.getRange(currentRow, 12).setValue(0.8); // 80% target
    sheet.getRange(currentRow, 10, 1, 3).setBackground(currentRow % 2 === 0 ? "#F5F5F5" : METRICS_BG_COLOR);
    
    // Format as percentage
    sheet.getRange(currentRow, 11, 1, 2).setNumberFormat("0.0%");
    
    // Add conditional formatting (green if meeting target, red if not)
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(sheet.getRange(currentRow, 12).getValue())
      .setBackground("#FFCDD2") // Light red if above target
      .setRanges([sheet.getRange(currentRow, 11)])
      .build();
    
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
    
    currentRow++;
  }
  
  // Add a border below the metrics table
  sheet.getRange(metricsStartRow + 1, 10, currentRow - metricsStartRow - 1, 3).setBorder(
    true, true, true, true, true, true, 
    "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Add Expense Categories table with improved formatting
  let expenseStartRow = currentRow + 2;
  
  // Add Expense Categories header
  sheet.getRange(expenseStartRow, 10).setValue("Expense Categories");
  sheet.getRange(expenseStartRow, 10, 1, 6)
    .setBackground(HEADER_BG_COLOR)
    .setFontWeight("bold")
    .setFontColor(HEADER_TEXT_COLOR)
    .setHorizontalAlignment("center");
  
  expenseStartRow++;
  
  // Add headers
  sheet.getRange(expenseStartRow, 10).setValue("Expense");
  sheet.getRange(expenseStartRow, 11).setValue("Amount");
  sheet.getRange(expenseStartRow, 12).setValue("Rate");
  sheet.getRange(expenseStartRow, 13).setValue("Target Rate");
  sheet.getRange(expenseStartRow, 14).setValue("% change");
  sheet.getRange(expenseStartRow, 15).setValue("Amount change");
  
  sheet.getRange(expenseStartRow, 10, 1, 6)
    .setBackground("#F5F5F5")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  // Find expense categories (these should be subcategories under "Expenses")
  const expenseCategories = [];
  
  // Look for expense categories based on the user's specific expense types
  const expenseTypes = ["Essentials", "Wants/Pleasure", "Extra"];
  
  for (let i = 1; i < data.length; i++) {
    // Check if this row has a type that's considered an expense
    if (expenseTypes.includes(data[i][0]) && data[i][1]) {
      expenseCategories.push({
        category: data[i][1],
        type: data[i][0],
        row: i + 1
      });
    }
  }
  
  // Add rows for each expense category with improved formatting
  currentRow = expenseStartRow + 1;
  expenseCategories.forEach(category => {
    sheet.getRange(currentRow, 10).setValue(category.category);
    
    // Use the Average column (Q) for more accurate calculations
    sheet.getRange(currentRow, 11).setFormula(`=Q${category.row}`);
    sheet.getRange(currentRow, 12).setFormula(`=IFERROR(K${currentRow}/Q${incomeRow}, 0)`);
    
    // Set target rate based on expense type
    let targetRate = 0.2; // Default
    if (category.type === "Essentials") {
      targetRate = 0.5; // 50% for essentials
    } else if (category.type === "Wants/Pleasure") {
      targetRate = 0.3; // 30% for wants
    } else if (category.type === "Extra") {
      targetRate = 0.2; // 20% for extras
    }
    
    sheet.getRange(currentRow, 13).setValue(targetRate);
    sheet.getRange(currentRow, 14).setFormula(`=IFERROR((L${currentRow}-M${currentRow})/M${currentRow}, 0)`);
    sheet.getRange(currentRow, 15).setFormula(`=IFERROR(K${currentRow}-(Q${incomeRow}*M${currentRow}), 0)`);
    
    // Apply alternating row colors
    sheet.getRange(currentRow, 10, 1, 6).setBackground(currentRow % 2 === 0 ? "#F5F5F5" : METRICS_BG_COLOR);
    
    // Format cells
    formatAsCurrency(sheet.getRange(currentRow, 11)); // Amount column as currency
    sheet.getRange(currentRow, 12, 1, 1).setNumberFormat("0.0%"); // Rate column as percentage
    sheet.getRange(currentRow, 13, 1, 1).setNumberFormat("0.0%"); // Target Rate column as percentage
    sheet.getRange(currentRow, 14, 1, 1).setNumberFormat("0.0%"); // % change column as percentage
    formatAsCurrency(sheet.getRange(currentRow, 15)); // Amount change column as currency
    
    // Add conditional formatting for the % change column
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#FFCDD2") // Light red if over budget
      .setRanges([sheet.getRange(currentRow, 14)])
      .build();
    
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
    
    currentRow++;
  });
  
  // Add Total Expenses row with distinct formatting
  sheet.getRange(currentRow, 10).setValue("Total Expenses");
  sheet.getRange(currentRow, 11).setFormula(`=Q${expensesRow}`);
  sheet.getRange(currentRow, 12).setFormula(`=IFERROR(K${currentRow}/Q${incomeRow}, 0)`);
  sheet.getRange(currentRow, 13).setValue(1); // Target 100%
  sheet.getRange(currentRow, 14).setFormula(`=IFERROR((L${currentRow}-M${currentRow})/M${currentRow}, 0)`);
  sheet.getRange(currentRow, 15).setFormula(`=IFERROR(K${currentRow}-(Q${incomeRow}*M${currentRow}), 0)`);
  
  // Format the total row
  sheet.getRange(currentRow, 10, 1, 6)
    .setBackground("#C62828")
    .setFontWeight("bold")
    .setFontColor("#FFFFFF");
  
  // Format cells
  formatAsCurrency(sheet.getRange(currentRow, 11));
  sheet.getRange(currentRow, 12, 1, 2).setNumberFormat("0.0%");
  sheet.getRange(currentRow, 14).setNumberFormat("0.0%");
  formatAsCurrency(sheet.getRange(currentRow, 15));
  
  // Add borders to the expense table
  sheet.getRange(expenseStartRow, 10, currentRow - expenseStartRow + 1, 6).setBorder(
    true, true, true, true, true, true, 
    "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Create expenditure breakdown chart only if we have expense categories
  if (expenseCategories.length > 0) {
    createExpenditureChart(sheet, expenseStartRow + 1, currentRow - 1, 10);
  }
}

/**
 * Creates an enhanced chart for expenditure breakdown
 */
function createExpenditureChart(sheet, startRow, endRow, categoryCol) {
  // Define color constants for better consistency
  const CHART_COLORS = [
    "#C62828", // Red (for Essentials)
    "#FF8F00", // Amber (for Wants/Pleasure)
    "#1565C0", // Blue (for Extra)
    "#2E7D32", // Green
    "#6A1B9A", // Purple
    "#E64A19", // Deep Orange
    "#00695C", // Teal
    "#5D4037"  // Brown
  ];
  
  // Define chart data range (category name and amount)
  const dataRange = sheet.getRange(startRow, categoryCol, endRow - startRow + 1, 2);
  
  // Create a pie chart with enhanced styling
  const chartBuilder = sheet.newChart();
  chartBuilder
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setPosition(startRow, categoryCol + 6, 0, 0)
    .setOption('title', 'Expenditure Breakdown')
    .setOption('titleTextStyle', {
      color: '#424242',
      fontSize: 16,
      bold: true
    })
    .setOption('pieSliceText', 'percentage')
    .setOption('pieHole', 0.4) // Create a donut chart for more modern look
    .setOption('legend', { 
      position: 'right',
      textStyle: {
        color: '#424242',
        fontSize: 12
      }
    })
    .setOption('colors', CHART_COLORS)
    .setOption('width', 450)
    .setOption('height', 300)
    .setOption('is3D', false)
    .setOption('pieSliceTextStyle', {
      color: '#FFFFFF',
      fontSize: 14,
      bold: true
    })
    .setOption('tooltip', { 
      showColorCode: true,
      textStyle: { fontSize: 12 }
    });
  
  // Add the chart to the sheet
  sheet.insertChart(chartBuilder.build());
  
  // Create a second chart - a column chart showing expense categories vs target
  const columnChartBuilder = sheet.newChart();
  
  // Define data range for the column chart (category, amount, target amount)
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
      color: '#424242',
      fontSize: 16,
      bold: true
    })
    .setOption('legend', { 
      position: 'top',
      textStyle: {
        color: '#424242',
        fontSize: 12
      }
    })
    .setOption('colors', ['#C62828', '#2E7D32']) // Red for actual, green for target
    .setOption('width', 450)
    .setOption('height', 300)
    .setOption('hAxis', {
      title: 'Category',
      titleTextStyle: {color: '#424242'},
      textStyle: {color: '#424242', fontSize: 10}
    })
    .setOption('vAxis', {
      title: 'Rate (% of Income)',
      titleTextStyle: {color: '#424242'},
      textStyle: {color: '#424242'},
      format: 'percent'
    })
    .setOption('bar', {groupWidth: '75%'})
    .setOption('isStacked', false);
  
  // Add the column chart to the sheet
  sheet.insertChart(columnChartBuilder.build());
}

/**
 * Formats the overview sheet for better readability
 */
function formatOverviewSheet(sheet) {
  // Define color constants for better consistency
  const INCOME_COLOR = "#388E3C"; // Green for income values
  const EXPENSE_COLOR = "#D32F2F"; // Red for expense values
  const SAVINGS_COLOR = "#1565C0"; // Blue for savings values
  const NEUTRAL_COLOR = "#424242"; // Dark gray for neutral values
  const BORDER_COLOR = "#FF8F00"; // Amber for borders
  
  // Format currency columns (adjusted for Shared? column)
  const lastRow = sheet.getLastRow();
  const currencyColumns = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]; // All month columns and Average (E to R)
  
  currencyColumns.forEach(col => {
    formatAsCurrency(sheet.getRange(2, col, lastRow - 1, 1));
  });
  
  // We'll format percentage columns only in their specific sections, not globally
  
  // Adjust column widths for better readability
  sheet.setColumnWidth(1, 150); // Type
  sheet.setColumnWidth(2, 150); // Category
  sheet.setColumnWidth(3, 150); // Sub-Category
  sheet.setColumnWidth(4, 80);  // Shared?
  
  // Set month column widths to be consistent
  for (let i = 5; i <= 16; i++) {
    sheet.setColumnWidth(i, 90); // Month columns
  }
  
  sheet.setColumnWidth(17, 100); // Average column
  
  // Format the metrics section columns
  sheet.setColumnWidth(10, 150); // Expense category
  sheet.setColumnWidth(11, 100); // Amount
  sheet.setColumnWidth(12, 80);  // Rate
  sheet.setColumnWidth(13, 80);  // Target Rate
  sheet.setColumnWidth(14, 80);  // % change
  sheet.setColumnWidth(15, 100); // Amount change
  
  // Add borders between major sections
  const data = sheet.getDataRange().getValues();
  
  // Find rows containing total sections
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].startsWith("Total ")) {
      // Add a bottom border to total rows
      sheet.getRange(i + 1, 1, 1, 17).setBorder(
        null, null, true, null, null, null, 
        BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    }
  }
  
  // Format total rows with bold text and distinct background
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].startsWith("Total ")) {
      const row = i + 1;
      const totalType = data[i][0].replace("Total ", "");
      let bgColor;
      let fontColor = "#FFFFFF"; // White text for all total rows
      
      if (totalType === "Income") {
        bgColor = "#2E7D32"; // Green for Income
      } else if (totalType === "Expenses") {
        bgColor = "#C62828"; // Red for Expenses
      } else if (totalType === "Savings") {
        bgColor = "#1565C0"; // Blue for Savings
      } else {
        bgColor = "#424242"; // Dark gray for other types
      }
      
      sheet.getRange(row, 1, 1, 17)
        .setBackground(bgColor)
        .setFontWeight("bold")
        .setFontColor(fontColor);
    }
  }
  
  // Format the Net calculation rows
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].startsWith("Net (")) {
      const row = i + 1;
      sheet.getRange(row, 1, 1, 17)
        .setBackground("#424242")
        .setFontWeight("bold")
        .setFontColor("#FFFFFF");
      
      // Format the values in the Net row
      for (let col = 5; col <= 17; col++) {
        // Apply conditional formatting based on value
        const cell = sheet.getRange(row, col);
        formatAsCurrency(cell);
      }
    }
  }
  
  // Apply number formatting to all value cells
  for (let row = 2; row <= lastRow; row++) {
    for (let col = 5; col <= 17; col++) {
      const cell = sheet.getRange(row, col);
      
      // Check if this is a total or net row (already formatted above)
      const rowValue = data[row - 1][0] || "";
      if (rowValue.startsWith("Total ") || rowValue.startsWith("Net (")) {
        continue;
      }
      
      // Apply conditional formatting based on the row type
      const rowType = data[row - 1][0] || "";
      if (rowType === "Income") {
        formatAsCurrency(cell);
        cell.setFontColor(INCOME_COLOR);
      } else if (rowType === "Expenses") {
        formatAsCurrency(cell);
        cell.setFontColor(EXPENSE_COLOR);
      } else if (rowType === "Savings") {
        formatAsCurrency(cell);
        cell.setFontColor(SAVINGS_COLOR);
      }
    }
  }
  
  // No alternating row colors as per new requirements
  // We're keeping rows with no background color except for headers and totals
}
