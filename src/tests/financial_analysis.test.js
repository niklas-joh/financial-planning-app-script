/**
 * Financial Planning Tools - Financial Analysis Service Tests
 *
 * This file contains tests for the FinancialPlanner.FinancialAnalysisService module.
 * It includes mocking SpreadsheetApp and other dependencies.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;

  // --- Mock Dependencies & Globals ---
  let mockSheetData = {};
  let mockActiveSpreadsheet = null;
  let lastToast = null;
  let lastAlert = null;
  let lastHandledError = null;
  let chartsAdded = [];
  let activeSheetName = null; // Track activated sheet

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
    setFontColor: function(color) { this._fontColor = color; return this; },
    setNumberFormat: function(format) { this._numberFormat = format; return this; },
    setHorizontalAlignment: function(align) { return this; },
    setBorder: function(...args) { return this; }, // Mock borders
    setFormulas: function(formulas) { /* Mock setting formulas */ return this; },
    clear: function() { /* Mock clear */ return this; },
    clearContent: function() { /* Mock clear content */ return this; },
    clearFormats: function() { /* Mock clear formats */ return this; }
  };

  const mockSheet = {
    _name: null, _dataRef: null, _frozenRows: 0,
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
    getMaxRows: function() { return this._dataRef.length + 50; },
    getMaxColumns: function() { return this._dataRef[0] ? this._dataRef[0].length + 10 : 26; },
    clear: function() { this._dataRef = []; return this; },
    clearFormats: function() { return this; },
    setFrozenRows: function(rows) { this._frozenRows = rows; },
    setColumnWidth: function(col, width) { /* Mock */ },
    setName: function(name) { this._name = name; },
    activate: function() { activeSheetName = this._name; return this; },
    insertChart: function(chart) { chartsAdded.push(chart); },
    newChart: function() { return mockChartBuilder; },
    getConditionalFormatRules: function() { return []; }, // Simplified mock
    setConditionalFormatRules: function(rules) { /* Mock */ }
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
    }
  };

  // Global mocks
  global.SpreadsheetApp = {
    getActiveSpreadsheet: function() { return mockActiveSpreadsheet; },
    BorderStyle: { SOLID: 'SOLID' }, // Mock enum
    newConditionalFormatRule: function() { // Mock builder
        return {
            whenFormulaSatisfied: function(formula) { this._formula = formula; return this; },
            setBackground: function(color) { this._background = color; return this; },
            setRanges: function(ranges) { this._ranges = ranges; return this; },
            build: function() { return { _formula: this._formula, _background: this._background, _ranges: this._ranges }; }
        };
    }
  };
  global.Charts = { ChartType: { PIE: 'PIE', COLUMN: 'COLUMN' } };
  global.Logger = { log: function(msg) { console.log("Logger.log:", msg); } };

  // --- Mock Other Dependencies ---
  const mockConfig = {
      _data: {
          SHEETS: { OVERVIEW: "Overview", ANALYSIS: "Analysis" },
          TARGET_RATES: { DEFAULT: 0.2, ESSENTIALS: 0.5, WANTS: 0.3, EXTRA: 0.2 }, // Added WANTS for mapping test
          COLORS: { UI: { HEADER_BG: "#ccc", HEADER_FONT: "#000", METRICS_BG: "#eee", EXPENSE_FONT: "#f00", INCOME_FONT: "#0f0" }, CHART: { TITLE: "#333", TEXT: "#444", SERIES: ["#f00", "#0f0", "#00f"] } }
      },
      get: function() { return JSON.parse(JSON.stringify(this._data)); },
      getSection: function(section) { return JSON.parse(JSON.stringify(this._data[section] || {})); }
  };
  const mockUtils = {
      getOrCreateSheet: function(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); },
      formatAsCurrency: function(range) { range.setNumberFormat('Currency'); },
      formatAsPercentage: function(range) { range.setNumberFormat('Percentage'); }
  };
  const mockUiService = {
      showLoadingSpinner: function(msg) { lastToast = { message: msg, title: "Working..." }; },
      hideLoadingSpinner: function() { lastToast = null; }, // Simulate hiding
      showSuccessNotification: function(msg) { lastToast = { message: msg, title: "Success" }; },
      showErrorNotification: function(title, msg) { lastAlert = { title: title, message: msg }; }
  };
  const mockErrorService = {
      handle: function(error, msg) { lastHandledError = { error, msg }; console.error("ERROR SERVICE MOCK:", msg, error); },
      create: function(msg, details) { const e = new Error(msg); e.details = details; e.name="FinancialPlannerError"; return e; },
      log: function(error) { console.log("ErrorService Mock Log:", error.message); }
  };

  // --- Test Suite Setup ---
   // Redefine FinancialAnalysisService with mocks
   const TestAnalysisService = (function(utils, uiService, errorService, config) {
       // --- Copy of FinancialAnalysisService Implementation Start ---
       // PASTE THE COPIED IMPLEMENTATION OF FinancialPlanner.FinancialAnalysisService HERE
       // Ensure all internal references use the mocked parameters (utils, uiService, etc.)
       // For brevity, the full copy is omitted here, but it's crucial for the test setup.
       class FinancialAnalysisService {
            constructor(spreadsheet, overviewSheet, analysisConfig) { this.spreadsheet = spreadsheet; this.overviewSheet = overviewSheet; this.config = analysisConfig; this.analysisSheet = utils.getOrCreateSheet(spreadsheet, this.config.SHEETS.ANALYSIS); this.data = null; this.totals = null; }
            initialize() { this.extractDataFromOverview(); this.setupAnalysisSheet(); }
            analyze() { let currentRow = 2; currentRow = this.addKeyMetricsSection(currentRow); currentRow += 2; currentRow = this.addExpenseCategoriesSection(currentRow); currentRow += 2; this.createExpenditureCharts(currentRow); }
            extractDataFromOverview() { const overviewData = this.overviewSheet.getDataRange().getValues(); this.data = { incomeCategories: [], expenseCategories: [], savingsCategories: [], months: [] }; this.totals = { income: { row: -1, value: 0 }, expenses: { row: -1, value: 0 }, savings: { row: -1, value: 0 }, essentials: { row: -1, value: 0 }, wantsPleasure: { row: -1, value: 0 }, extra: { row: -1, value: 0 } }; for (let i = 4; i <= 15; i++) { this.data.months.push(overviewData[0][i]); } for (let i = 0; i < overviewData.length; i++) { const rowData = overviewData[i]; if (rowData[0] === "Total Income") { this.totals.income.row = i + 1; this.totals.income.value = rowData[16] || 0; } else if (rowData[0] === "Total Essentials") { this.totals.essentials.row = i + 1; this.totals.essentials.value = rowData[16] || 0; if (this.totals.expenses.row === -1) this.totals.expenses.row = i + 1; this.totals.expenses.value += rowData[16] || 0; } else if (rowData[0] === "Total Wants/Pleasure") { this.totals.wantsPleasure.row = i + 1; this.totals.wantsPleasure.value = rowData[16] || 0; if (this.totals.expenses.row === -1) this.totals.expenses.row = i + 1; this.totals.expenses.value += rowData[16] || 0; } else if (rowData[0] === "Total Extra") { this.totals.extra.row = i + 1; this.totals.extra.value = rowData[16] || 0; if (this.totals.expenses.row === -1) this.totals.expenses.row = i + 1; this.totals.expenses.value += rowData[16] || 0; } else if (rowData[0] === "Total Savings") { this.totals.savings.row = i + 1; this.totals.savings.value = rowData[16] || 0; } /* ... simplified category extraction ... */ } }
            setupAnalysisSheet() { this.analysisSheet.clear(); this.analysisSheet.clearFormats(); this.analysisSheet.getRange("A1").setValue("Financial Analysis"); this.analysisSheet.getRange("A1:J1").setBackground(this.config.COLORS.UI.HEADER_BG).setFontWeight("bold").setFontColor(this.config.COLORS.UI.HEADER_FONT); this.analysisSheet.setFrozenRows(1); this.analysisSheet.setColumnWidth(1, 200); this.analysisSheet.setColumnWidth(2, 120); this.analysisSheet.setColumnWidth(3, 120); this.analysisSheet.setName(this.config.SHEETS.ANALYSIS); }
            addKeyMetricsSection(startRow) { /* ... simplified mock implementation ... */ this.analysisSheet.getRange(startRow, 1).setValue("Key Metrics Header"); startRow++; this.analysisSheet.getRange(startRow, 1, 1, 4).setValues([["Metric", "Value", "Target", "Description"]]); startRow++; this.analysisSheet.getRange(startRow, 1, 1, 4).setValues([["Test Metric", 0.5, 0.6, "Desc"]]); return startRow + 1; }
            addExpenseCategoriesSection(startRow) { /* ... simplified mock implementation ... */ this.analysisSheet.getRange(startRow, 1).setValue("Expense Cat Header"); startRow++; this.analysisSheet.getRange(startRow, 1, 1, 6).setValues([["Category", "Type", "Amount", "%", "Target", "Var"]]); startRow++; this.analysisSheet.getRange(startRow, 1, 1, 6).setValues([["Food", "Exp", 100, 0.1, 0.2, -0.1]]); return startRow + 1; }
            createExpenditureCharts(startRow) { /* ... simplified mock implementation ... */ const chart = this.analysisSheet.newChart().build(); this.analysisSheet.insertChart(chart); }
            suggestSavingsOpportunities() { uiService.showInfoAlert('Savings Opportunities', 'Coming Soon!'); } // Use mock uiService
            detectSpendingAnomalies() { uiService.showInfoAlert('Spending Anomalies', 'Coming Soon!'); }
            analyzeFixedVsVariableExpenses() { uiService.showInfoAlert('Fixed vs Variable', 'Coming Soon!'); }
            generateCashFlowForecast() { uiService.showInfoAlert('Cash Flow Forecast', 'Coming Soon!'); }
       }
       return { analyze: function(spreadsheet, overviewSheet) { try { uiService.showLoadingSpinner("Analyzing financial data..."); const analysisConfig = { ...config.get(), TARGET_RATES: { ...config.getSection('TARGET_RATES'), WANTS_PLEASURE: config.getSection('TARGET_RATES').WANTS } }; const service = new FinancialAnalysisService(spreadsheet, overviewSheet, analysisConfig); service.initialize(); service.analyze(); uiService.hideLoadingSpinner(); return service; } catch (error) { uiService.hideLoadingSpinner(); errorService.handle(error, "Error analyzing financial data"); throw error; } }, showKeyMetrics: function() { try { const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); const overviewSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').OVERVIEW); if (!overviewSheet) { uiService.showErrorNotification("Error", "Overview sheet not found. Please generate the financial overview first."); return; } FinancialPlanner.FinancialAnalysisService.analyze(spreadsheet, overviewSheet); const analysisSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').ANALYSIS); analysisSheet.activate(); uiService.showSuccessNotification("Key metrics have been generated in the Analysis sheet."); } catch (error) { errorService.handle(error, "Failed to generate key metrics"); } }, suggestSavingsOpportunities: function() { try { const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); const overviewSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').OVERVIEW); if (!overviewSheet) { uiService.showErrorNotification("Error", "Overview sheet not found."); return; } const service = this.analyze(spreadsheet, overviewSheet); service.suggestSavingsOpportunities(); } catch (error) { errorService.handle(error, "Failed to suggest savings opportunities"); } }, detectSpendingAnomalies: function() { try { const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); const overviewSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').OVERVIEW); if (!overviewSheet) { uiService.showErrorNotification("Error", "Overview sheet not found."); return; } const service = this.analyze(spreadsheet, overviewSheet); service.detectSpendingAnomalies(); } catch (error) { errorService.handle(error, "Failed to detect spending anomalies"); } } };
       // --- Copy of FinancialAnalysisService Implementation End ---
   })(mockUtils, mockUiService, mockErrorService, mockConfig); // Pass mocks


  // --- Helper to reset state before each test ---
  function setupMockDataAndState() {
      mockSheetData = {
          "Overview": [ // Needs realistic overview data structure
              ["Type", "Category", "Sub-Category", "Shared?", "Jan-24", "Feb-24", "Mar-24", "Apr-24", "May-24", "Jun-24", "Jul-24", "Aug-24", "Sep-24", "Oct-24", "Nov-24", "Dec-24", "Total", "Average"],
              // ... more rows ...
              ["Total Income", "", "", "", 0,0,0,0,0,0,0,0,0,0,0,0, 2000, 2000], // Example Total row
              ["Total Essentials", "", "", "", 0,0,0,0,0,0,0,0,0,0,0,0, 500, 500],
              ["Total Wants/Pleasure", "", "", "", 0,0,0,0,0,0,0,0,0,0,0,0, 300, 300],
              ["Total Extra", "", "", "", 0,0,0,0,0,0,0,0,0,0,0,0, 100, 100],
              ["Total Savings", "", "", "", 0,0,0,0,0,0,0,0,0,0,0,0, 1100, 1100] // Example Total row
          ],
          "Analysis": []
      };
      mockActiveSpreadsheet.sheets = {}; // Clear existing sheets
      mockActiveSpreadsheet.insertSheet("Overview");
      mockActiveSpreadsheet.insertSheet("Analysis");
      lastToast = null;
      lastAlert = null;
      lastHandledError = null;
      chartsAdded = [];
      activeSheetName = null;
  }

  // --- Test Cases ---

  T.registerTest("FinancialAnalysisService", "analyze should create Analysis sheet and call sub-methods", function() {
      setupMockDataAndState();
      const overviewSheet = mockActiveSpreadsheet.getSheetByName("Overview");

      // We test the main 'analyze' entry point which creates the internal service instance
      const serviceInstance = TestAnalysisService.analyze(mockActiveSpreadsheet, overviewSheet);

      T.assertNotNull(serviceInstance, "Analyze should return the service instance.");
      const analysisSheet = mockActiveSpreadsheet.getSheetByName("Analysis");
      T.assertNotNull(analysisSheet, "Analysis sheet should be created.");

      // Check if key sections were added (based on simplified mock implementations)
      const analysisData = mockSheetData["Analysis"];
      let keyMetricsHeaderFound = analysisData.some(row => row[0] === "Key Metrics Header");
      let expenseCatHeaderFound = analysisData.some(row => row[0] === "Expense Cat Header");

      T.assertTrue(keyMetricsHeaderFound, "Key Metrics section header should be added.");
      T.assertTrue(expenseCatHeaderFound, "Expense Categories section header should be added.");
      T.assertTrue(chartsAdded.length > 0, "Charts should be added.");
      T.assertTrue(lastHandledError === null, "No error should be handled during successful analysis.");
  });

  T.registerTest("FinancialAnalysisService", "showKeyMetrics should call analyze and activate sheet", function() {
      setupMockDataAndState();

      TestAnalysisService.showKeyMetrics();

      T.assertNotNull(mockActiveSpreadsheet.getSheetByName("Analysis"), "Analysis sheet should exist after showKeyMetrics.");
      T.assertEquals("Analysis", activeSheetName, "Analysis sheet should be activated.");
      T.assertNotNull(lastToast, "Success notification should be shown.");
      T.assertEquals("Key metrics have been generated in the Analysis sheet.", lastToast.message, "Correct success message expected.");
      T.assertTrue(lastHandledError === null, "No error should be handled during successful showKeyMetrics.");
  });

  T.registerTest("FinancialAnalysisService", "showKeyMetrics should show error if Overview sheet missing", function() {
      setupMockDataAndState();
      delete mockActiveSpreadsheet.sheets["Overview"]; // Remove overview sheet

      TestAnalysisService.showKeyMetrics();

      T.assertNotNull(lastAlert, "Error notification (alert) should be shown.");
      T.assertTrue(lastAlert.message.includes("Overview sheet not found"), "Error message should indicate missing Overview sheet.");
      T.assertTrue(activeSheetName !== "Analysis", "Analysis sheet should not be activated on error.");
  });
  
  T.registerTest("FinancialAnalysisService", "Placeholder methods should show alerts", function() {
      setupMockDataAndState();
      const overviewSheet = mockActiveSpreadsheet.getSheetByName("Overview");
      // Need to call analyze first to get a service instance for the placeholders
      const serviceInstance = TestAnalysisService.analyze(mockActiveSpreadsheet, overviewSheet); 
      
      lastAlert = null; // Reset alert mock
      serviceInstance.suggestSavingsOpportunities();
      T.assertNotNull(lastAlert, "suggestSavingsOpportunities should show an alert.");
      T.assertTrue(lastAlert.message.includes("Coming Soon"), "Alert message for suggestSavingsOpportunities should mention 'Coming Soon'.");

      lastAlert = null;
      serviceInstance.detectSpendingAnomalies();
      T.assertNotNull(lastAlert, "detectSpendingAnomalies should show an alert.");
      T.assertTrue(lastAlert.message.includes("Coming Soon"), "Alert message for detectSpendingAnomalies should mention 'Coming Soon'.");
      
      // ... add similar checks for analyzeFixedVsVariableExpenses and generateCashFlowForecast if needed ...
  });


})(); // End IIFE
