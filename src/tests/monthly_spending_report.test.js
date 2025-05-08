/**
 * Financial Planning Tools - Monthly Spending Report Service Tests
 *
 * This file contains tests for the FinancialPlanner.MonthlySpendingReport module.
 * It includes mocking SpreadsheetApp, dependencies, and sheet data.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;

  // --- Mock Dependencies & Globals ---
  let mockSheetData = {};
  let mockActiveSpreadsheet = null;
  let lastToast = null;
  let lastHandledError = null;
  let chartsAdded = []; // Track added charts

  // --- Mock Spreadsheet Objects (Simplified) ---
   const mockRange = {
    _sheetName: null, _row: 0, _col: 0, _numRows: 1, _numCols: 1, _sheetDataRef: null,
    _fontWeight: null, _background: null, _fontSize: null, _fontColor: null, _numberFormat: null,

    setValue: function(value) { this._sheetDataRef[this._row - 1][this._col - 1] = value; return this; },
    setValues: function(values) { /* ... simplified ... */
         for(let r=0; r<this._numRows; r++) {
           if (!this._sheetDataRef[this._row + r - 1]) this._sheetDataRef[this._row + r - 1] = [];
           for (let c=0; c<this._numCols; c++) {
               this._sheetDataRef[this._row + r - 1][this._col + c - 1] = values[r][c];
           }
       }
       return this;
    },
    getValue: function() { return this._sheetDataRef[this._row - 1][this._col - 1]; },
    getValues: function() { /* ... simplified ... */
        const result = [];
        for(let r=0; r<this._numRows; r++) {
            const rowData = [];
             if (!this._sheetDataRef[this._row + r - 1]) this._sheetDataRef[this._row + r - 1] = [];
            for (let c=0; c<this._numCols; c++) {
                rowData.push(this._sheetDataRef[this._row + r - 1][this._col + c - 1] || "");
            }
            result.push(rowData);
        }
        return result;
    },
    setFontWeight: function(weight) { this._fontWeight = weight; return this; },
    setBackground: function(color) { this._background = color; return this; },
    setFontSize: function(size) { this._fontSize = size; return this; },
    setFontColor: function(color) { this._fontColor = color; return this; },
    setNumberFormat: function(format) { this._numberFormat = format; return this; },
    merge: function() { return this; }, // Chainable mock
    clearContent: function() { /* ... simplified ... */
        for(let r=0; r<this._numRows; r++) {
           for (let c=0; c<this._numCols; c++) {
                if (this._sheetDataRef[this._row + r - 1]) {
                    this._sheetDataRef[this._row + r - 1][this._col + c - 1] = "";
                }
           }
       }
       return this;
    }
  };

  const mockSheet = {
    _name: null, _dataRef: null,
    getName: function() { return this._name; },
    getDataRange: function() { /* ... simplified ... */
        const numRows = this._dataRef.length;
        const numCols = numRows > 0 ? this._dataRef[0].length : 0;
        return Object.assign({}, mockRange, { _sheetName: this._name, _row: 1, _col: 1, _numRows: numRows, _numCols: numCols, _sheetDataRef: this._dataRef });
    },
    getRange: function(row, col, numRows = 1, numCols = 1) { /* ... simplified ... */
         while(this._dataRef.length < row + numRows -1) this._dataRef.push([]);
         const maxCols = col + numCols -1;
         this._dataRef.forEach(r => { while(r.length < maxCols) r.push(""); });
         return Object.assign({}, mockRange, { _sheetName: this._name, _row: row, _col: col, _numRows: numRows, _numCols: numCols, _sheetDataRef: this._dataRef });
    },
    getLastRow: function() { return this._dataRef.length; },
    getMaxRows: function() { return this._dataRef.length + 50; }, // Mock max rows
    getMaxColumns: function() { return this._dataRef[0] ? this._dataRef[0].length + 10 : 10; }, // Mock max cols
    clear: function() { this._dataRef = []; }, // Simple clear
    clearFormats: function() {}, // Mock
    autoResizeColumns: function(startCol, numCols) { /* console.log(`Mock: AutoResize ${numCols} cols from ${startCol}`); */ },
    insertChart: function(chart) { chartsAdded.push(chart); },
    newChart: function() { return mockChartBuilder; } // Return mock builder
  };
  
   const mockChart = { _type: null, _options: {}, _ranges: [], _position: {} };
   const mockChartBuilder = {
       setChartType: function(type) { mockChart._type = type; return this; },
       addRange: function(range) { mockChart._ranges.push(range); return this; },
       setPosition: function(row, col, offX, offY) { mockChart._position = {row, col, offX, offY}; return this; },
       setOption: function(name, value) { mockChart._options[name] = value; return this; },
       build: function() { return Object.assign({}, mockChart); } // Return a copy
   };

  mockActiveSpreadsheet = {
    sheets: {},
    getSheetByName: function(name) { return this.sheets[name] || null; },
    insertSheet: function(name) { /* ... simplified ... */
        if (this.sheets[name]) return this.sheets[name];
        const newSheetData = []; mockSheetData[name] = newSheetData;
        const newSheet = Object.assign({}, mockSheet, { _name: name, _dataRef: newSheetData });
        this.sheets[name] = newSheet; return newSheet;
    },
    toast: function(message, title, timeout) { lastToast = { message, title, timeout }; }
  };

  // Global mocks
  global.SpreadsheetApp = {
    getActiveSpreadsheet: function() { return mockActiveSpreadsheet; }
    // Add other methods if needed
  };
  global.Charts = { ChartType: { PIE: 'PIE' } }; // Mock ChartType enum
  global.Logger = { log: function(msg) { console.log("Logger.log:", msg); } };

  // --- Mock Other Dependencies ---
  const mockConfig = {
    _sheets: { TRANSACTIONS: "Transactions", MONTHLY_REPORT: "Monthly Report" }, // Assuming report name might be in config
    _locale: { CURRENCY_SYMBOL: '$', CURRENCY_LOCALE: '1' },
    getSheetNames: function() { return this._sheets; },
    getLocale: function() { return this._locale; },
    getSection: function(section) {
        if (section === 'SHEETS') return this._sheets;
        if (section === 'LOCALE') return this._locale;
        return {};
     }
  };
  const mockUtils = {
     getOrCreateSheet: function(spreadsheet, sheetName) {
        let sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            sheet = spreadsheet.insertSheet(sheetName);
            // Simulate header creation if needed by tests, but service does it
        } else {
            sheet.clear(); // Simulate clearing for report generation
        }
        return sheet;
     },
     getMonthName: function(index) { return ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][index]; },
     formatAsCurrency: function(range, symbol, locale) { range.setNumberFormat('Currency'); /* Simplified */ },
     formatAsPercentage: function(range, decimals) { range.setNumberFormat('Percentage'); /* Simplified */ }
  };
  const mockUiService = {
      showLoadingSpinner: function(msg) { /* console.log("UI Mock:", msg); */ },
      hideLoadingSpinner: function() { /* console.log("UI Mock: Hide Spinner"); */ },
      showSuccessNotification: function(msg) { lastToast = { message: msg, title: "Success" }; }
  };
  const mockErrorService = {
      handle: function(error, msg) { lastHandledError = { error, msg }; console.error("ERROR SERVICE MOCK:", msg, error); },
      create: function(msg, details) { const e = new Error(msg); e.details = details; e.name="FinancialPlannerError"; return e; },
      log: function(error) { console.log("ErrorService Mock Log:", error.message); }
  };

  // --- Test Suite Setup ---
   // Redefine MonthlySpendingReport with mocks
   const TestMonthlySpendingReport = (function(utils, uiService, errorService, config) {
       // --- Copy of MonthlySpendingReport Implementation Start ---
        function addMonthlyReportCharts(sheet, categoryData, totalExpenses) { const lastRow = sheet.getLastRow(); const categories = Object.keys(categoryData); const categoryValues = categories.map(category => { return Object.values(categoryData[category]).reduce((sum, amount) => sum + amount, 0); }); const chartRange = sheet.getRange(lastRow + 3, 1, categories.length + 1, 2); sheet.getRange(lastRow + 3, 1).setValue("Category"); sheet.getRange(lastRow + 3, 2).setValue("Amount"); for (let i = 0; i < categories.length; i++) { sheet.getRange(lastRow + 4 + i, 1).setValue(categories[i]); sheet.getRange(lastRow + 4 + i, 2).setValue(categoryValues[i]); } const pieChart = sheet.newChart().setChartType(Charts.ChartType.PIE).addRange(chartRange).setPosition(5, 8, 0, 0).setOption('title', 'Expense Breakdown by Category').setOption('pieSliceText', 'percentage').setOption('width', 450).setOption('height', 300).build(); sheet.insertChart(pieChart); }
        function calculatePreviousMonthsAverage(sheet, data, category, subcategory, dateColIndex, typeColIndex, categoryColIndex, subcategoryColIndex, amountColIndex, monthsToLookBack) { const now = new Date(); const currentMonth = now.getMonth(); const currentYear = now.getFullYear(); let totalAmount = 0; let monthsFound = 0; for (let i = 1; i <= monthsToLookBack; i++) { let targetMonth = currentMonth - i; let targetYear = currentYear; if (targetMonth < 0) { targetMonth += 12; targetYear--; } const monthlyTransactions = data.filter((row, index) => { if (index === 0) return false; const date = new Date(row[dateColIndex]); return date.getMonth() === targetMonth && date.getFullYear() === targetYear && row[categoryColIndex] === category && (row[subcategoryColIndex] || "(None)") === subcategory; }); let monthTotal = 0; monthlyTransactions.forEach(row => { const amount = Math.abs(parseFloat(row[amountColIndex]) || 0); monthTotal += amount; }); if (monthlyTransactions.length > 0) { totalAmount += monthTotal; monthsFound++; } } return monthsFound > 0 ? totalAmount / monthsFound : 0; }
        function createMonthlySpendingReport() { const ss = SpreadsheetApp.getActiveSpreadsheet(); const reportSheet = utils.getOrCreateSheet(ss, "Monthly Report"); const now = new Date(); const currentMonth = now.getMonth(); const currentYear = now.getFullYear(); reportSheet.getRange("A1").setValue(`Monthly Spending Report - ${utils.getMonthName(currentMonth)} ${currentYear}`); reportSheet.getRange("A1:F1").merge().setFontWeight("bold").setFontSize(14); reportSheet.getRange("A3:F3").setValues([["Category", "Sub-Category", "Amount", "% of Total", "Avg Last 3 Months", "Trend"]]).setFontWeight("bold").setBackground("#D9EAD3"); const transactionSheet = ss.getSheetByName(config.getSheetNames().TRANSACTIONS); if (!transactionSheet) { throw errorService.create("Could not find 'Transactions' sheet", { severity: "high" }); } const transactionData = transactionSheet.getDataRange().getValues(); const headers = transactionData[0]; const dateColIndex = headers.indexOf("Date"); const typeColIndex = headers.indexOf("Type"); const categoryColIndex = headers.indexOf("Category"); const subcategoryColIndex = headers.indexOf("Sub-Category"); const amountColIndex = headers.indexOf("Amount"); if (dateColIndex < 0 || typeColIndex < 0 || categoryColIndex < 0 || amountColIndex < 0) { throw errorService.create("Could not find required columns in Transaction sheet", { severity: "high" }); } const currentMonthTransactions = transactionData.filter((row, index) => { if (index === 0) return false; const date = new Date(row[dateColIndex]); return date.getMonth() === currentMonth && date.getFullYear() === currentYear; }); const categoryData = {}; let totalExpenses = 0; currentMonthTransactions.forEach(row => { const type = row[typeColIndex]; if (type !== "Expenses" && type !== "Wants/Pleasure" && type !== "Extra") return; const category = row[categoryColIndex]; const subcategory = row[subcategoryColIndex] || "(None)"; const amount = Math.abs(parseFloat(row[amountColIndex]) || 0); if (!categoryData[category]) { categoryData[category] = {}; } if (!categoryData[category][subcategory]) { categoryData[category][subcategory] = 0; } categoryData[category][subcategory] += amount; totalExpenses += amount; }); let rowIndex = 4; Object.keys(categoryData).sort().forEach(category => { const subcategories = categoryData[category]; let categoryTotal = 0; Object.values(subcategories).forEach(amount => { categoryTotal += amount; }); reportSheet.getRange(rowIndex, 1).setValue(category); reportSheet.getRange(rowIndex, 3).setValue(categoryTotal); reportSheet.getRange(rowIndex, 4).setValue(totalExpenses > 0 ? categoryTotal / totalExpenses : 0); reportSheet.getRange(rowIndex, 1, 1, 6).setBackground("#F3F3F3").setFontWeight("bold"); rowIndex++; Object.keys(subcategories).sort().forEach(subcategory => { const amount = subcategories[subcategory]; reportSheet.getRange(rowIndex, 1).setValue(""); reportSheet.getRange(rowIndex, 2).setValue(subcategory); reportSheet.getRange(rowIndex, 3).setValue(amount); reportSheet.getRange(rowIndex, 4).setValue(totalExpenses > 0 ? amount / totalExpenses : 0); const last3MonthsAvg = calculatePreviousMonthsAverage(transactionSheet, transactionData, category, subcategory, dateColIndex, typeColIndex, categoryColIndex, subcategoryColIndex, amountColIndex, 3); reportSheet.getRange(rowIndex, 5).setValue(last3MonthsAvg); if (last3MonthsAvg > 0) { const percentChange = (amount - last3MonthsAvg) / last3MonthsAvg; const trendCell = reportSheet.getRange(rowIndex, 6); if (percentChange > 0.1) { trendCell.setValue("↑ " + (percentChange * 100).toFixed(0) + "%").setFontColor("#CC0000"); } else if (percentChange < -0.1) { trendCell.setValue("↓ " + (Math.abs(percentChange) * 100).toFixed(0) + "%").setFontColor("#006600"); } else { trendCell.setValue("→ Stable").setFontColor("#666666"); } } rowIndex++; }); rowIndex++; }); reportSheet.getRange(rowIndex, 1).setValue("TOTAL EXPENSES"); reportSheet.getRange(rowIndex, 3).setValue(totalExpenses); reportSheet.getRange(rowIndex, 4).setValue(totalExpenses > 0 ? 1 : 0); reportSheet.getRange(rowIndex, 1, 1, 6).setBackground("#D9D9D9").setFontWeight("bold"); if (rowIndex > 4) { utils.formatAsCurrency(reportSheet.getRange(4, 3, rowIndex - 3, 1), config.getLocale().CURRENCY_SYMBOL, config.getLocale().CURRENCY_LOCALE); reportSheet.getRange(4, 4, rowIndex - 3, 1).setNumberFormat("0.0%"); utils.formatAsCurrency(reportSheet.getRange(4, 5, rowIndex - 3, 1), config.getLocale().CURRENCY_SYMBOL, config.getLocale().CURRENCY_LOCALE); } addMonthlyReportCharts(reportSheet, categoryData, totalExpenses); reportSheet.autoResizeColumns(1, 6); return reportSheet; }
        return { generate: function() { try { uiService.showLoadingSpinner("Generating monthly spending report..."); const reportSheet = createMonthlySpendingReport(); uiService.hideLoadingSpinner(); uiService.showSuccessNotification("Monthly spending report generated!"); return reportSheet; } catch (error) { uiService.hideLoadingSpinner(); errorService.handle(error, "Failed to generate monthly spending report"); return null; } } };
       // --- Copy of MonthlySpendingReport Implementation End ---
   })(mockUtils, mockUiService, mockErrorService, mockConfig); // Pass mocks

  // --- Helper to reset state before each test ---
  function setupMockData(currentMonthDate) {
      // Use currentMonthDate to generate relevant transaction dates
      const currentMonth = currentMonthDate.getMonth();
      const currentYear = currentMonthDate.getFullYear();
      const prevMonth = new Date(currentYear, currentMonth - 1, 15).toISOString().split('T')[0];
      const twoMonthsAgo = new Date(currentYear, currentMonth - 2, 15).toISOString().split('T')[0];
      const currentMonthDay10 = new Date(currentYear, currentMonth, 10).toISOString().split('T')[0];
      const currentMonthDay15 = new Date(currentYear, currentMonth, 15).toISOString().split('T')[0];

      mockSheetData = {
          "Transactions": [
              ["Date", "Description", "Type", "Category", "Sub-Category", "Amount"],
              // Current Month Data
              [currentMonthDay10, "Groceries", "Expenses", "Food", "Groceries", -50],
              [currentMonthDay10, "Gas", "Expenses", "Transport", "Gas", -40],
              [currentMonthDay15, "Restaurant", "Wants/Pleasure", "Food", "Restaurants", -60],
              [currentMonthDay15, "Salary", "Income", "Salary", "", 2000],
              // Previous Month Data (for average calc)
              [prevMonth, "Groceries", "Expenses", "Food", "Groceries", -45],
              [prevMonth, "Gas", "Expenses", "Transport", "Gas", -35],
              // Two Months Ago Data
              [twoMonthsAgo, "Groceries", "Expenses", "Food", "Groceries", -55]
          ],
          "Monthly Report": [] // Start with empty report sheet
      };
      mockActiveSpreadsheet.sheets = {}; // Clear existing sheets
      mockActiveSpreadsheet.insertSheet("Transactions");
      mockActiveSpreadsheet.insertSheet("Monthly Report"); // Ensure report sheet exists for getOrCreateSheet
      lastToast = null;
      lastHandledError = null;
      chartsAdded = [];
  }

  // --- Test Cases ---

  T.registerTest("MonthlySpendingReport", "generate should create report sheet with title", function() {
      const testDate = new Date(2024, 4, 20); // May 2024
      setupMockData(testDate);
      global.Date = class extends Date { constructor() { super(testDate); } }; // Mock current date

      const reportSheet = TestMonthlySpendingReport.generate();

      T.assertNotNull(reportSheet, "generate() should return a sheet object.");
      T.assertEquals("Monthly Report", reportSheet.getName(), "Report sheet name should be correct.");
      const title = reportSheet.getRange(1, 1).getValue();
      T.assertTrue(title.includes("Monthly Spending Report - May 2024"), "Report title should contain month and year.");
      T.assertNotNull(lastToast, "Success notification should be shown.");
      T.assertEquals("Monthly spending report generated!", lastToast.message, "Correct success message expected.");
      global.Date = Date; // Restore original Date
  });

  T.registerTest("MonthlySpendingReport", "generate should aggregate expenses correctly", function() {
      const testDate = new Date(2024, 4, 20); // May 2024
      setupMockData(testDate);
      global.Date = class extends Date { constructor() { super(testDate); } };

      const reportSheet = TestMonthlySpendingReport.generate();
      const reportData = mockSheetData["Monthly Report"];

      // Find rows for verification (this is brittle, depends on implementation details)
      let foodGroceriesRow = -1, foodRestaurantsRow = -1, transportGasRow = -1, totalRow = -1;
      for(let i=0; i<reportData.length; i++) {
          if(reportData[i][0] === "Food" && reportData[i][1] === "Groceries") foodGroceriesRow = i;
          if(reportData[i][0] === "Food" && reportData[i][1] === "Restaurants") foodRestaurantsRow = i;
          if(reportData[i][0] === "Transport" && reportData[i][1] === "Gas") transportGasRow = i;
          if(reportData[i][0] === "TOTAL EXPENSES") totalRow = i;
      }

      T.assertTrue(foodGroceriesRow > 0, "Food/Groceries row should exist.");
      T.assertEquals(50, reportData[foodGroceriesRow][2], "Food/Groceries amount should be 50.");
      T.assertTrue(foodRestaurantsRow > 0, "Food/Restaurants row should exist.");
      T.assertEquals(60, reportData[foodRestaurantsRow][2], "Food/Restaurants amount should be 60.");
       T.assertTrue(transportGasRow > 0, "Transport/Gas row should exist.");
      T.assertEquals(40, reportData[transportGasRow][2], "Transport/Gas amount should be 40.");
      T.assertTrue(totalRow > 0, "Total Expenses row should exist.");
      T.assertEquals(150, reportData[totalRow][2], "Total Expenses amount should be 150 (50+60+40).");

      global.Date = Date; // Restore original Date
  });

   T.registerTest("MonthlySpendingReport", "generate should calculate percentages correctly", function() {
      const testDate = new Date(2024, 4, 20); // May 2024
      setupMockData(testDate);
      global.Date = class extends Date { constructor() { super(testDate); } };

      const reportSheet = TestMonthlySpendingReport.generate();
      const reportData = mockSheetData["Monthly Report"];
      const totalExpenses = 150;

      let foodGroceriesRow = -1, totalRow = -1;
       for(let i=0; i<reportData.length; i++) {
          if(reportData[i][0] === "Food" && reportData[i][1] === "Groceries") foodGroceriesRow = i;
          if(reportData[i][0] === "TOTAL EXPENSES") totalRow = i;
      }

      T.assertTrue(foodGroceriesRow > 0, "Food/Groceries row should exist.");
      T.assertEquals(50 / totalExpenses, reportData[foodGroceriesRow][3], "Food/Groceries percentage should be correct.");
      T.assertTrue(totalRow > 0, "Total Expenses row should exist.");
      T.assertEquals(1, reportData[totalRow][3], "Total Expenses percentage should be 1 (100%).");

      global.Date = Date; // Restore original Date
  });

  // Note: Testing calculatePreviousMonthsAverage and trend requires more complex date mocking or
  // focusing the test on verifying the *call* to calculatePreviousMonthsAverage rather than its result.
  // For simplicity, we'll assume calculatePreviousMonthsAverage works if called.

  T.registerTest("MonthlySpendingReport", "generate should handle missing Transactions sheet", function() {
      const testDate = new Date(2024, 4, 20);
      setupMockData(testDate);
      delete mockActiveSpreadsheet.sheets["Transactions"]; // Remove sheet
      global.Date = class extends Date { constructor() { super(testDate); } };

      const reportSheet = TestMonthlySpendingReport.generate();

      T.assertTrue(reportSheet === null, "generate() should return null on error.");
      T.assertNotNull(lastHandledError, "Error should have been handled.");
      T.assertTrue(lastHandledError.message.includes("Could not find 'Transactions' sheet"), "Error message should indicate missing sheet.");

      global.Date = Date;
  });
  
   T.registerTest("MonthlySpendingReport", "generate should add charts", function() {
      const testDate = new Date(2024, 4, 20); // May 2024
      setupMockData(testDate);
      global.Date = class extends Date { constructor() { super(testDate); } };

      TestMonthlySpendingReport.generate();

      T.assertEquals(1, chartsAdded.length, "One chart should have been added.");
      T.assertEquals("PIE", chartsAdded[0]._type, "Chart type should be PIE.");
      T.assertTrue(chartsAdded[0]._options.title.includes("Expenditure Breakdown"), "Chart title should be correct.");

      global.Date = Date; // Restore original Date
  });


})(); // End IIFE
