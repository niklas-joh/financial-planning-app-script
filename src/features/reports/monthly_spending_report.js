/**
 * Financial Planning Tools - Monthly Spending Report Module
 * 
 * This file provides functionality for generating monthly spending reports.
 * It follows the namespace pattern and uses dependency injection for better maintainability.
 */

/**
 * @namespace FinancialPlanner.MonthlySpendingReport
 * @description Service for generating a detailed monthly spending report sheet.
 * It analyzes transactions for the current month, categorizes expenses, calculates averages,
 * identifies trends, and adds visualizations.
 * @param {FinancialPlanner.Utils} utils - The utility service.
 * @param {FinancialPlanner.UIService} uiService - The UI service for notifications.
 * @param {FinancialPlanner.ErrorService} errorService - The error handling service.
 * @param {FinancialPlanner.Config} config - The global configuration service.
 */
FinancialPlanner.MonthlySpendingReport = (function(utils, uiService, errorService, config) {
  // Private variables and functions
  
  /**
   * Adds relevant charts (e.g., pie chart for expense breakdown) to the monthly report sheet.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object where the report is generated.
   * @param {object<string, object<string, number>>} categoryData - Processed data grouped by category and sub-category with summed amounts.
   * @param {number} totalExpenses - The total expense amount for the month.
   * @return {void}
   * @private
   */
  function addMonthlyReportCharts(sheet, categoryData, totalExpenses) {
    // Find last row with data
    const lastRow = sheet.getLastRow();
    
    // Create pie chart for category breakdown
    const categories = Object.keys(categoryData);
    const categoryValues = categories.map(category => {
      return Object.values(categoryData[category]).reduce((sum, amount) => sum + amount, 0);
    });
    
    // Create temporary range for chart data
    const chartDataRangeStartRow = lastRow + 3;
    const numChartDataRows = categories.length + 1; // +1 for header
    const chartRange = sheet.getRange(chartDataRangeStartRow, 1, numChartDataRows, 2);
    
    // Prepare data for batch write
    const chartData = [["Category", "Amount"]]; // Header row
    for (let i = 0; i < categories.length; i++) {
      chartData.push([categories[i], categoryValues[i]]);
    }
    
    // Batch write chart data
    chartRange.setValues(chartData);
    
    // Create the chart using the same range
    const pieChart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(chartRange)
      .setPosition(5, 8, 0, 0)
      .setOption('title', 'Expense Breakdown by Category')
      .setOption('pieSliceText', 'percentage')
      .setOption('width', 450)
      .setOption('height', 300)
      .build();
    
    sheet.insertChart(pieChart);
  }
  
  /**
   * Calculates the average monthly spending for a specific category and sub-category
   * over a defined number of previous months.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The transaction sheet object (currently unused but passed).
   * @param {Array<Array<any>>} data - The raw transaction data (2D array).
   * @param {string} category - The category to filter by.
   * @param {string} subcategory - The sub-category to filter by (use "(None)" if no sub-category).
   * @param {number} dateColIndex - 0-based index of the 'Date' column.
   * @param {number} typeColIndex - 0-based index of the 'Type' column.
   * @param {number} categoryColIndex - 0-based index of the 'Category' column.
   * @param {number} subcategoryColIndex - 0-based index of the 'Sub-Category' column.
   * @param {number} amountColIndex - 0-based index of the 'Amount' column.
   * @param {number} monthsToLookBack - The number of previous months to include in the average calculation.
   * @return {number} The calculated average monthly spending for the specified criteria, or 0 if no data found.
   * @private
   */
  function calculatePreviousMonthsAverage(
    sheet, 
    data, 
    category, 
    subcategory, 
    dateColIndex, 
    typeColIndex, 
    categoryColIndex, 
    subcategoryColIndex,
    amountColIndex, 
    monthsToLookBack
  ) {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    let totalAmount = 0;
    let monthsFound = 0;
    
    // Look back for the specified number of months
    for (let i = 1; i <= monthsToLookBack; i++) {
      let targetMonth = currentMonth - i;
      let targetYear = currentYear;
      
      // Handle year change
      if (targetMonth < 0) {
        targetMonth += 12;
        targetYear--;
      }
      
      // Filter transactions for the target month
      const monthlyTransactions = data.filter((row, index) => {
        if (index === 0) return false; // Skip header
        
        const date = new Date(row[dateColIndex]);
        return date.getMonth() === targetMonth && 
               date.getFullYear() === targetYear &&
               row[categoryColIndex] === category &&
               (row[subcategoryColIndex] || "(None)") === subcategory;
      });
      
      // Sum amounts for this month
      let monthTotal = 0;
      monthlyTransactions.forEach(row => {
        const amount = Math.abs(parseFloat(row[amountColIndex]) || 0);
        monthTotal += amount;
      });
      
      if (monthlyTransactions.length > 0) {
        totalAmount += monthTotal;
        monthsFound++;
      }
    }
    
    // Return average or 0 if no data found
    return monthsFound > 0 ? totalAmount / monthsFound : 0;
  }
  
  /**
   * Core function to generate the monthly spending report sheet.
   * It fetches transaction data, filters for the current month's expenses,
   * groups data, calculates totals and averages, formats the sheet, and adds charts.
   * @return {GoogleAppsScript.Spreadsheet.Sheet} The generated or updated report sheet object.
   * @throws {FinancialPlannerError} If the 'Transactions' sheet or required columns are not found.
   * @private
   */
  function createMonthlySpendingReport() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get or create the Monthly Report sheet
    const reportSheet = utils.getOrCreateSheet(ss, "Monthly Report");
    
    // Get current month and year
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    // Set title
    reportSheet.getRange("A1").setValue(`Monthly Spending Report - ${utils.getMonthName(currentMonth)} ${currentYear}`);
    reportSheet.getRange("A1:F1").merge().setFontWeight("bold").setFontSize(14);
    
    // Headers
    reportSheet.getRange("A3").setValue("Category");
    reportSheet.getRange("B3").setValue("Sub-Category");
    reportSheet.getRange("C3").setValue("Amount");
    reportSheet.getRange("D3").setValue("% of Total");
    reportSheet.getRange("E3").setValue("Avg Last 3 Months");
    reportSheet.getRange("F3").setValue("Trend");
    
    reportSheet.getRange("A3:F3").setFontWeight("bold").setBackground("#D9EAD3");
    
    // Get transaction data
    const transactionSheet = ss.getSheetByName(config.getSheetNames().TRANSACTIONS);
    if (!transactionSheet) {
      throw errorService.create("Could not find 'Transactions' sheet", { severity: "high" });
    }
    
    const transactionData = transactionSheet.getDataRange().getValues();
    const headers = transactionData[0];
    
    // Find column indices
    const dateColIndex = headers.indexOf("Date");
    const typeColIndex = headers.indexOf("Type");
    const categoryColIndex = headers.indexOf("Category");
    const subcategoryColIndex = headers.indexOf("Sub-Category");
    const amountColIndex = headers.indexOf("Amount");
    
    if (dateColIndex < 0 || typeColIndex < 0 || categoryColIndex < 0 || amountColIndex < 0) {
      throw errorService.create("Could not find required columns in Transaction sheet", { severity: "high" });
    }
    
    // Filter transactions for current month
    const currentMonthTransactions = transactionData.filter((row, index) => {
      if (index === 0) return false; // Skip header
      
      const date = new Date(row[dateColIndex]);
      return date.getMonth() === currentMonth && date.getFullYear() === currentYear;
    });
    
    // Group by category and sub-category
    const categoryData = {};
    let totalExpenses = 0;
    
    currentMonthTransactions.forEach(row => {
      const type = row[typeColIndex];
      // Only include expenses
      if (type !== "Expenses" && type !== "Wants/Pleasure" && type !== "Extra") return;
      
      const category = row[categoryColIndex];
      const subcategory = row[subcategoryColIndex] || "(None)";
      const amount = Math.abs(parseFloat(row[amountColIndex]) || 0);
      
      if (!categoryData[category]) {
        categoryData[category] = {};
      }
      
      if (!categoryData[category][subcategory]) {
        categoryData[category][subcategory] = 0;
      }
      
      categoryData[category][subcategory] += amount;
      totalExpenses += amount;
    });

    // --- Batch Write Data ---
    const reportData = []; // Array to hold all row data for batch write
    const formatInfo = []; // Array to hold info for formatting after batch write
    let currentRowIndex = 4; // Start after headers

    Object.keys(categoryData).sort().forEach(category => {
      const subcategories = categoryData[category];
      let categoryTotal = 0;
      Object.values(subcategories).forEach(amount => { categoryTotal += amount; });

      // Prepare category header row data
      const categoryRowData = [category, "", categoryTotal, totalExpenses > 0 ? categoryTotal / totalExpenses : 0, "", ""];
      reportData.push(categoryRowData);
      formatInfo.push({ row: currentRowIndex, type: 'categoryHeader' });
      currentRowIndex++;

      // Prepare subcategory row data
      Object.keys(subcategories).sort().forEach(subcategory => {
        const amount = subcategories[subcategory];
        const last3MonthsAvg = calculatePreviousMonthsAverage(
          transactionSheet, transactionData, category, subcategory,
          dateColIndex, typeColIndex, categoryColIndex, subcategoryColIndex, amountColIndex, 3
        );
        
        let trendValue = "";
        let trendColor = null;
        if (last3MonthsAvg > 0) {
          const percentChange = (amount - last3MonthsAvg) / last3MonthsAvg;
          if (percentChange > 0.1) {
            trendValue = "↑ " + (percentChange * 100).toFixed(0) + "%";
            trendColor = "#CC0000"; // Red
          } else if (percentChange < -0.1) {
            trendValue = "↓ " + (Math.abs(percentChange) * 100).toFixed(0) + "%";
            trendColor = "#006600"; // Green
          } else {
            trendValue = "→ Stable";
            trendColor = "#666666"; // Gray
          }
        }

        const subCategoryRowData = ["", subcategory, amount, totalExpenses > 0 ? amount / totalExpenses : 0, last3MonthsAvg, trendValue];
        reportData.push(subCategoryRowData);
        formatInfo.push({ row: currentRowIndex, type: 'subcategory', trendColor: trendColor });
        currentRowIndex++;
      });

      // Add empty row for spacing (optional, can be handled by formatting later)
      reportData.push(["", "", "", "", "", ""]);
      formatInfo.push({ row: currentRowIndex, type: 'spacer' });
      currentRowIndex++;
    });

    // Prepare total row data
    const totalRowData = ["TOTAL EXPENSES", "", totalExpenses, totalExpenses > 0 ? 1 : 0, "", ""];
    reportData.push(totalRowData);
    formatInfo.push({ row: currentRowIndex, type: 'totalRow' });
    const finalDataRowIndex = currentRowIndex; // Keep track of the last data row

    // Batch write all data
    if (reportData.length > 0) {
        reportSheet.getRange(4, 1, reportData.length, 6).setValues(reportData);
    }

    // --- Apply Formatting (can still involve loops, but fewer I/O calls) ---
    
    // Apply number formats in batches
    if (finalDataRowIndex >= 4) {
        // Amount column (C) and Avg column (E)
        utils.formatAsCurrency(reportSheet.getRange(4, 3, finalDataRowIndex - 4 + 1, 1), config.getLocale().CURRENCY_SYMBOL, config.getLocale().CURRENCY_LOCALE);
        utils.formatAsCurrency(reportSheet.getRange(4, 5, finalDataRowIndex - 4 + 1, 1), config.getLocale().CURRENCY_SYMBOL, config.getLocale().CURRENCY_LOCALE);
        // Percentage column (D)
        reportSheet.getRange(4, 4, finalDataRowIndex - 4 + 1, 1).setNumberFormat("0.0%");
    }

    // Apply row-specific formatting based on collected info
    formatInfo.forEach(info => {
        const range = reportSheet.getRange(info.row, 1, 1, 6);
        if (info.type === 'categoryHeader') {
            range.setBackground("#F3F3F3").setFontWeight("bold");
        } else if (info.type === 'totalRow') {
            range.setBackground("#D9D9D9").setFontWeight("bold");
        } else if (info.type === 'subcategory' && info.trendColor) {
            reportSheet.getRange(info.row, 6).setFontColor(info.trendColor);
        }
        // Note: Spacers don't need specific formatting here unless desired
    });
    
    // Add charts
    addMonthlyReportCharts(reportSheet, categoryData, totalExpenses);
    
    // Auto-size columns
    reportSheet.autoResizeColumns(1, 6);
    
    return reportSheet;
  }
  
  // Public API
  return {
    /**
     * Public method to generate the monthly spending report.
     * Wraps the private `createMonthlySpendingReport` function with UI feedback and error handling.
     * @return {GoogleAppsScript.Spreadsheet.Sheet | null} The generated report sheet object, or null if an error occurred.
     * @public
     * @example
     * // Called from a menu item or controller:
     * FinancialPlanner.MonthlySpendingReport.generate();
     */
    generate: function() {
      try {
        uiService.showLoadingSpinner("Generating monthly spending report...");
        const reportSheet = createMonthlySpendingReport();
        uiService.hideLoadingSpinner();
        uiService.showSuccessNotification("Monthly spending report generated!");
        return reportSheet;
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Failed to generate monthly spending report");
        return null;
      }
    }
  };
})(FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.ErrorService, FinancialPlanner.Config);

// Backward compatibility layer for existing global functions
/**
 * Generates the monthly spending report.
 * Maintained for backward compatibility with older triggers or direct calls.
 * Delegates to `FinancialPlanner.MonthlySpendingReport.generate()`.
 * @return {GoogleAppsScript.Spreadsheet.Sheet | null | undefined} The report sheet object, null if an error occurred during generation,
 *         or undefined if the service isn't loaded.
 * @global
 */
function generateMonthlySpendingReport() {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.MonthlySpendingReport && FinancialPlanner.MonthlySpendingReport.generate) {
    return FinancialPlanner.MonthlySpendingReport.generate();
  }
  Logger.log("Global generateMonthlySpendingReport: FinancialPlanner.MonthlySpendingReport not available.");
  // Optionally show an error to the user if appropriate for a direct call scenario
  // SpreadsheetApp.getUi().alert("Error: Monthly Spending Report module not loaded.");
}
