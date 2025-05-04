/**
 * Creates a monthly spending report sheet
 * This function analyzes current month transactions and compares with previous months
 */
function generateMonthlySpendingReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create the Monthly Report sheet
  const reportSheet = getOrCreateSheet(ss, "Monthly Report");
  
  // Get current month and year
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  
  // Set title
  reportSheet.getRange("A1").setValue(`Monthly Spending Report - ${getMonthName(currentMonth)} ${currentYear}`);
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
  const transactionSheet = ss.getSheetByName("Transactions");
  if (!transactionSheet) {
    SpreadsheetApp.getUi().alert("Error: Could not find 'Transactions' sheet");
    return;
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
    SpreadsheetApp.getUi().alert("Error: Could not find required columns in Transaction sheet");
    return;
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
  
  // Write data to report sheet
  let rowIndex = 4;
  
  Object.keys(categoryData).sort().forEach(category => {
    const subcategories = categoryData[category];
    let categoryTotal = 0;
    
    // Calculate category total
    Object.values(subcategories).forEach(amount => {
      categoryTotal += amount;
    });
    
    // Add category header
    reportSheet.getRange(rowIndex, 1).setValue(category);
    reportSheet.getRange(rowIndex, 3).setValue(categoryTotal);
    reportSheet.getRange(rowIndex, 4).setValue(categoryTotal / totalExpenses);
    
    // Format category row
    reportSheet.getRange(rowIndex, 1, 1, 6).setBackground("#F3F3F3").setFontWeight("bold");
    
    rowIndex++;
    
    // Add subcategories
    Object.keys(subcategories).sort().forEach(subcategory => {
      const amount = subcategories[subcategory];
      
      reportSheet.getRange(rowIndex, 1).setValue(""); // Empty category
      reportSheet.getRange(rowIndex, 2).setValue(subcategory);
      reportSheet.getRange(rowIndex, 3).setValue(amount);
      reportSheet.getRange(rowIndex, 4).setValue(amount / totalExpenses);
      
      // Calculate average for last 3 months (excluding current)
      const last3MonthsAvg = calculatePreviousMonthsAverage(
        transactionSheet, 
        transactionData,
        category,
        subcategory,
        dateColIndex,
        typeColIndex,
        categoryColIndex,
        subcategoryColIndex,
        amountColIndex,
        3
      );
      
      reportSheet.getRange(rowIndex, 5).setValue(last3MonthsAvg);
      
      // Add trend indicator
      if (last3MonthsAvg > 0) {
        const percentChange = (amount - last3MonthsAvg) / last3MonthsAvg;
        
        if (percentChange > 0.1) {
          reportSheet.getRange(rowIndex, 6).setValue("↑ " + (percentChange * 100).toFixed(0) + "%");
          reportSheet.getRange(rowIndex, 6).setFontColor("#CC0000"); // Red for increase
        } else if (percentChange < -0.1) {
          reportSheet.getRange(rowIndex, 6).setValue("↓ " + (Math.abs(percentChange) * 100).toFixed(0) + "%");
          reportSheet.getRange(rowIndex, 6).setFontColor("#006600"); // Green for decrease
        } else {
          reportSheet.getRange(rowIndex, 6).setValue("→ Stable");
          reportSheet.getRange(rowIndex, 6).setFontColor("#666666"); // Gray for stable
        }
      }
      
      rowIndex++;
    });
    
    rowIndex++; // Add space between categories
  });
  
  // Add total row
  reportSheet.getRange(rowIndex, 1).setValue("TOTAL EXPENSES");
  reportSheet.getRange(rowIndex, 3).setValue(totalExpenses);
  reportSheet.getRange(rowIndex, 4).setValue(1); // 100%
  
  reportSheet.getRange(rowIndex, 1, 1, 6).setBackground("#D9D9D9").setFontWeight("bold");
  
  // Format columns
  reportSheet.getRange(4, 3, rowIndex - 3, 1).setNumberFormat("€#,##0.00");
  reportSheet.getRange(4, 4, rowIndex - 3, 1).setNumberFormat("0.0%");
  reportSheet.getRange(4, 5, rowIndex - 3, 1).setNumberFormat("€#,##0.00");
  
  // Add charts
  addMonthlyReportCharts(reportSheet, categoryData, totalExpenses);
  
  // Auto-size columns
  reportSheet.autoResizeColumns(1, 6);
  
  // Notify user
  SpreadsheetApp.getActiveSpreadsheet().toast("Monthly spending report generated!", "Success");
}

/**
 * Creates charts for the monthly spending report
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
  const chartRange = sheet.getRange(lastRow + 3, 1, categories.length + 1, 2);
  
  // Add header
  sheet.getRange(lastRow + 3, 1).setValue("Category");
  sheet.getRange(lastRow + 3, 2).setValue("Amount");
  
  // Add data
  for (let i = 0; i < categories.length; i++) {
    sheet.getRange(lastRow + 4 + i, 1).setValue(categories[i]);
    sheet.getRange(lastRow + 4 + i, 2).setValue(categoryValues[i]);
  }
  
  // Create the chart
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
 * Calculates the average spending for a specific category/subcategory over previous months
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
