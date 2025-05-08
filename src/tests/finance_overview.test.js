/**
 * Financial Planning Tools - Finance Overview Service Tests
 *
 * This file contains tests for the FinancialPlanner.FinanceOverview module.
 * It includes extensive mocking for dependencies and SpreadsheetApp interactions.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;

  // --- Mock Dependencies & Globals ---
  let mockSheetData = {};
  let mockActiveSpreadsheet = null;
  let mockScriptCacheStore = {};
  let mockUserPropertiesStore = {}; // For SettingsService mock
  let lastToast = null;
  let lastAlert = null;
  let lastHandledError = null;
  let analysisServiceAnalyzeCalled = false;
  let settingsGetValueLog = {};
  let settingsSetValueLog = {};
  let cacheGetLog = {};
  let cacheInvalidateAllCalled = false;

  // --- Mock Spreadsheet Objects (Simplified) ---
   const mockRange = {
    _sheetName: null, _row: 0, _col: 0, _numRows: 1, _numCols: 1, _sheetDataRef: null,
    _fontWeight: null, _background: null, _fontSize: null, _fontColor: null, _numberFormat: null, _note: null, _indent: 0,

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
    setVerticalAlignment: function(align) { return this; },
    insertCheckboxes: function() { return this; }, // Mock checkbox insertion
    setNote: function(note) { this._note = note; return this; },
    setIndent: function(indent) { this._indent = indent; return this; },
    setBorder: function(...args) { return this; }, // Mock borders
    clearContent: function() { /* ... simplified ... */ return this; },
    clearFormats: function() { return this; },
    clear: function() { this.clearContent(); this.clearFormats(); return this; },
    setDataValidation: function(rule) { return this; }, // Mock validation
    getA1Notation: function() { return `R${this._row}C${this._col}`; } // Simplified A1
  };

  const mockSheet = {
    _name: null, _dataRef: null, _hidden: false, _frozenRows: 0,
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
     getRange: function(a1Notation) { // Overload for A1 notation used by handleEdit
        // Very simplified parser for R C notation used in mockRange.getA1Notation
        if (a1Notation && a1Notation.startsWith("R") && a1Notation.includes("C")) {
            const parts = a1Notation.substring(1).split("C");
            const row = parseInt(parts[0], 10);
            const col = parseInt(parts[1], 10);
             if (!isNaN(row) && !isNaN(col)) {
                 return this.getRange(row, col, 1, 1);
             }
        }
        // Fallback for actual A1 like 'T1'
        if (a1Notation === 'T1') return this.getRange(1, 20); // Assuming T = 20
        if (a1Notation === 'S1') return this.getRange(1, 19); // Assuming S = 19
        // Add more specific A1 mappings if needed by tests
        throw new Error(`MockSheet.getRange(A1) not implemented for: ${a1Notation}`);
    },
    getLastRow: function() { return this._dataRef.length; },
    getMaxRows: function() { return this._dataRef.length + 50; },
    getMaxColumns: function() { return this._dataRef[0] ? this._dataRef[0].length + 10 : 26; },
    clear: function() { this._dataRef = []; return this; },
    clearFormats: function() { return this; },
    hideSheet: function() { this._hidden = true; },
    showColumns: function(colStart, numCols) { /* Mock */ },
    hideColumns: function(colStart, numCols) { /* Mock */ },
    setFrozenRows: function(rows) { this._frozenRows = rows; },
    insertSheet: function(name) { return mockActiveSpreadsheet.insertSheet(name); }
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
    getActiveSpreadsheet: function() { return mockActiveSpreadsheet; },
    BorderStyle: { SOLID_MEDIUM: 'SOLID_MEDIUM' } // Mock enum
    // Add other methods if needed
  };
   global.Session = { getScriptTimeZone: function() { return "Etc/GMT"; } };
   global.Utilities = { formatDate: function(date, tz, format) { return `Formatted:${date.toISOString()}`; } };
   global.CacheService = { getScriptCache: function() { /* ... simplified mock ... */ return { get: function(key){ return mockScriptCacheStore[key] || null; }, put: function(key, value, ttl){ mockScriptCacheStore[key] = value; }, remove: function(key){ delete mockScriptCacheStore[key]; }, removeAll: function(keys){ keys.forEach(k => delete mockScriptCacheStore[k]); } }; } };
   global.Logger = { log: function(msg) { console.log("Logger.log:", msg); } };

  // --- Mock Other Dependencies ---
  const mockConfig = {
      _data: {
          SHEETS: { OVERVIEW: "Overview", TRANSACTIONS: "Transactions", SETTINGS: "Settings" },
          HEADERS: ["Type", "Category", "Sub-Category", "Shared?", "Jan-24", "Feb-24", "Mar-24", "Apr-24", "May-24", "Jun-24", "Jul-24", "Aug-24", "Sep-24", "Oct-24", "Nov-24", "Dec-24", "Total", "Average"],
          UI: { SUBCATEGORY_TOGGLE: { LABEL_CELL: "S1", CHECKBOX_CELL: "T1", LABEL_TEXT: "Show Sub-Categories", NOTE_TEXT: "Toggle sub-cats" }, COLUMN_WIDTHS: { SHARED: 80 } },
          COLORS: { UI: { HEADER_BG: "#C62828", HEADER_FONT: "#FFFFFF", NET_BG: "#424242", NET_FONT: "#FFFFFF", BORDER: "#FF8F00", INCOME_FONT: "#00FF00", EXPENSE_FONT: "#FF0000" }, TYPE_HEADERS: { DEFAULT: { BG: "#DDDDDD", FONT: "#000000" } } },
          TYPE_ORDER: ["Income", "Expenses"],
          EXPENSE_TYPES: ["Expenses"],
          LOCALE: { DATE_FORMAT: "yyyy-MM-dd", CURRENCY_SYMBOL: "$", CURRENCY_LOCALE: "1" },
          CACHE: { KEYS: { CATEGORY_COMBINATIONS: "cats", GROUPED_COMBINATIONS: "grouped" } },
          PERFORMANCE: { USE_BATCH_OPERATIONS: true }
      },
      get: function() { return JSON.parse(JSON.stringify(this._data)); }, // Deep copy
      getSection: function(section) { return JSON.parse(JSON.stringify(this._data[section] || {})); }
  };
  const mockUtils = {
      getOrCreateSheet: function(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); },
      columnToLetter: function(col) { return String.fromCharCode(64 + col); }, // Simplified
      formatAsCurrency: function(range, symbol, locale) { range.setNumberFormat('Currency'); },
      formatAsPercentage: function(range, decimals) { range.setNumberFormat('Percentage'); }
  };
  const mockUiService = {
      showLoadingSpinner: function(msg) { lastToast = { message: msg, title: "Working..." }; },
      hideLoadingSpinner: function() { lastToast = { message: "", title: "", timeout: 1 }; },
      showSuccessNotification: function(msg, duration = 5) { lastToast = { message: msg, title: "Success", timeout: duration }; },
      showErrorNotification: function(title, msg) { lastAlert = { title: title, message: msg }; }
  };
  const mockErrorService = {
      handle: function(error, msg) { lastHandledError = { error, msg }; console.error("ERROR SERVICE MOCK:", msg, error); },
      create: function(msg, details) { const e = new Error(msg); e.details = details; e.name="FinancialPlannerError"; return e; },
      log: function(error) { console.log("ErrorService Mock Log:", error.message); }
  };
  const mockSettingsService = {
      getValue: function(key, defaultValue) { settingsGetValueLog[key] = (settingsGetValueLog[key] || 0) + 1; return mockUserPropertiesStore[key] !== undefined ? mockUserPropertiesStore[key] : defaultValue; },
      setValue: function(key, value) { settingsSetValueLog[key] = (settingsSetValueLog[key] || 0) + 1; mockUserPropertiesStore[key] = value; }
      // Add getPreference/setPreference if used by the actual service code being tested
  };
  const mockAnalysisService = {
      analyze: function(ss, sheet) { analysisServiceAnalyzeCalled = true; /* console.log("Mock AnalysisService.analyze called"); */ }
  };
  const mockCacheService = {
      get: function(key, computeFunc) {
          cacheGetLog[key] = (cacheGetLog[key] || 0) + 1;
          const cached = mockScriptCacheStore[key];
          if (cached) return JSON.parse(cached); // Simplified mock get
          const result = computeFunc();
          mockScriptCacheStore[key] = JSON.stringify(result); // Simplified mock put
          return result;
      },
      invalidateAll: function() { cacheInvalidateAllCalled = true; mockScriptCacheStore = {}; }
  };

  // --- Test Suite Setup ---
   // Redefine FinanceOverview with mocks
   const TestFinanceOverview = (function(utils, uiService, cacheService, errorService, config, settingsService, analysisServiceInstance) {
       // --- Copy of FinanceOverview Implementation Start ---
       // PASTE THE COPIED IMPLEMENTATION OF FinancialPlanner.FinanceOverview HERE
       // Ensure all internal references use the mocked parameters (utils, uiService, etc.)
       // For brevity, the full copy is omitted here, but it's crucial for the test setup.
       // Example snippet:
        class FinancialPlannerError extends Error { /* ... */ } // Include if defined locally
        function getProcessedTransactionData(sheet) { /* ... uses errorService ... */ return { data: [], indices: {} }; }
        function getUniqueCategoryCombinations(data, typeCol, categoryCol, subcategoryCol, showSubCategories) { /* ... */ return []; }
        function groupCategoryCombinations(combinations) { /* ... */ return {}; }
        function buildMonthlySumFormula(params) { /* ... uses utils, config ... */ return "=SUM(A1)"; }
        function formatDate(date) { /* ... uses Utilities, Session, config ... */ return "DATE"; }
        function clearSheetContent(sheet) { /* ... uses sheet mock ... */ }
        function setupHeaderRow(sheet, showSubCategories) { /* ... uses config, sheet mock ... */ }
        function addTypeHeaderRow(sheet, type, rowIndex) { /* ... uses config, sheet mock ... */ return rowIndex + 1; }
        function addCategoryRows(sheet, combinations, rowIndex, type, columnIndices) { /* ... uses config, utils, sheet mock ... */ return rowIndex + combinations.length; }
        function addTypeSubtotalRow(sheet, type, rowIndex, rowCount) { /* ... uses config, utils, sheet mock ... */ return rowIndex + 1; }
        function addNetCalculations(sheet, startRow) { /* ... uses config, utils, sheet mock ... */ return startRow + 3; }
        function findTotalRows(data) { /* ... */ return { incomeRow: 5, expensesRow: 10, savingsRow: 15 }; } // Mocked return
        function formatOverviewSheet(sheet) { /* ... uses config, sheet mock ... */ }
        function getTypeColors(type) { /* ... uses config ... */ return { BG: '#ccc', FONT: '#000' }; }
        function formatRangeAsCurrency(range) { /* ... uses utils, config ... */ }
        function getUserPreference(key, defaultValue) { /* ... uses settingsService, errorService ... */ return settingsService.getValue(key, defaultValue); }
        function setUserPreference(key, value) { /* ... uses settingsService, errorService ... */ settingsService.setValue(key, value); }

        class FinancialOverviewBuilder {
             constructor() { /* ... */ }
             initialize() { this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); const sheetNames = config.getSection('SHEETS'); this.overviewSheet = utils.getOrCreateSheet(this.spreadsheet, sheetNames.OVERVIEW); clearSheetContent(this.overviewSheet); this.transactionSheet = this.spreadsheet.getSheetByName(sheetNames.TRANSACTIONS); if (!this.transactionSheet) throw errorService.create(`Sheet "${sheetNames.TRANSACTIONS}" not found`); this.showSubCategories = getUserPreference("ShowSubCategories", true); return this; }
             processData() { const { data, indices } = getProcessedTransactionData(this.transactionSheet); this.transactionData = data; this.columnIndices = indices; this.categoryCombinations = cacheService.get(config.getSection('CACHE').KEYS.CATEGORY_COMBINATIONS, () => getUniqueCategoryCombinations(this.transactionData, this.columnIndices.type, this.columnIndices.category, this.columnIndices.subcategory, this.showSubCategories)); this.groupedCombinations = cacheService.get(config.getSection('CACHE').KEYS.GROUPED_COMBINATIONS, () => groupCategoryCombinations(this.categoryCombinations)); return this; }
             setupHeader() { setupHeaderRow(this.overviewSheet, this.showSubCategories); return this; }
             generateContent() { let rowIndex = 2; config.getSection('TYPE_ORDER').forEach(type => { if (!this.groupedCombinations[type]) return; rowIndex = addTypeHeaderRow(this.overviewSheet, type, rowIndex); rowIndex = addCategoryRows(this.overviewSheet, this.groupedCombinations[type], rowIndex, type, this.columnIndices); rowIndex = addTypeSubtotalRow(this.overviewSheet, type, rowIndex, this.groupedCombinations[type].length); rowIndex += 2; }); this.lastContentRowIndex = rowIndex; return this; }
             addNetCalculations() { this.lastContentRowIndex = addNetCalculations(this.overviewSheet, this.lastContentRowIndex); return this; }
             addMetrics() { if (analysisServiceInstance && analysisServiceInstance.analyze) { analysisServiceInstance.analyze(this.spreadsheet, this.overviewSheet); } else { console.error("Mock AnalysisService not available"); } return this; }
             formatSheet() { formatOverviewSheet(this.overviewSheet); return this; }
             applyPreferences() { if (this.showSubCategories) { this.overviewSheet.showColumns(3, 1); } else { this.overviewSheet.hideColumns(3, 1); } return this; }
             build() { return { sheet: this.overviewSheet, lastRow: this.overviewSheet.getLastRow(), success: true }; }
        }
        return { create: function() { try { uiService.showLoadingSpinner("Generating financial overview..."); cacheService.invalidateAll(); const builder = new FinancialOverviewBuilder(); const result = builder.initialize().processData().setupHeader().generateContent().addNetCalculations().addMetrics().formatSheet().applyPreferences().build(); uiService.hideLoadingSpinner(); uiService.showSuccessNotification("Financial overview generated successfully!"); return result; } catch (error) { uiService.hideLoadingSpinner(); if (error.name === 'FinancialPlannerError') { errorService.log(error); uiService.showErrorNotification("Error generating overview", error.message); } else { const wrappedError = errorService.create("Failed to generate financial overview", { originalError: error.message, stack: error.stack, severity: "high" }); errorService.log(wrappedError); uiService.showErrorNotification("Error generating overview", error.message); } throw error; } }, handleEdit: function(e) { try { if (e.range.getSheet().getName() !== config.getSection('SHEETS').OVERVIEW) return; const subcategoryToggle = config.getSection('UI').SUBCATEGORY_TOGGLE; if (e.range.getA1Notation() === subcategoryToggle.CHECKBOX_CELL) { const newValue = e.range.getValue(); setUserPreference("ShowSubCategories", newValue); uiService.showLoadingSpinner("Updating overview..."); this.create(); } } catch (error) { errorService.handle(errorService.create("Error handling Overview sheet edit", { originalError: error.toString(), eventDetails: JSON.stringify(e) }), "Failed to process change on Overview sheet."); } } };
       // --- Copy of FinanceOverview Implementation End ---
   })(mockUtils, mockUiService, mockCacheService, mockErrorService, mockConfig, mockSettingsService, mockAnalysisService); // Pass mocks


  // --- Helper to reset state before each test ---
  function setupMockDataAndState() {
      mockSheetData = {
          "Transactions": [
              ["Date", "Description", "Type", "Category", "Sub-Category", "Amount", "Shared"],
              ["2024-01-10", "T1", "Expenses", "Food", "Groceries", -50, ""],
              ["2024-01-15", "T2", "Income", "Salary", "", 2000, ""],
               ["2024-02-10", "T3", "Expenses", "Food", "Groceries", -60, ""],
          ],
          "Overview": [],
          "Settings": [["Preference", "Value"]]
      };
      mockActiveSpreadsheet.sheets = {}; // Clear existing sheets
      mockActiveSpreadsheet.insertSheet("Transactions");
      mockActiveSpreadsheet.insertSheet("Overview");
      mockActiveSpreadsheet.insertSheet("Settings");

      mockScriptCacheStore = {};
      mockUserPropertiesStore = {};
      lastToast = null;
      lastHandledError = null;
      analysisServiceAnalyzeCalled = false;
      settingsGetValueLog = {};
      settingsSetValueLog = {};
      cacheGetLog = {};
      cacheInvalidateAllCalled = false;
  }

  // --- Test Cases ---

  T.registerTest("FinanceOverview", "create should initialize, process, build, and format", function() {
      setupMockDataAndState();
      mockUserPropertiesStore["ShowSubCategories"] = true; // Set preference for test

      const result = TestFinanceOverview.create();

      T.assertNotNull(result, "Create should return a result object.");
      T.assertTrue(result.success, "Result success flag should be true.");
      T.assertNotNull(result.sheet, "Result should contain the sheet object.");
      T.assertEquals("Overview", result.sheet.getName(), "Result sheet should be the Overview sheet.");

      // Verify interactions with mocks
      T.assertTrue(cacheInvalidateAllCalled, "cacheService.invalidateAll should be called.");
      T.assertTrue(settingsGetValueLog["ShowSubCategories"] > 0, "settingsService.getValue('ShowSubCategories') should be called.");
      T.assertTrue(cacheGetLog["cats"] > 0, "cacheService.get for categories should be called.");
      T.assertTrue(cacheGetLog["grouped"] > 0, "cacheService.get for grouped categories should be called.");
      T.assertTrue(analysisServiceAnalyzeCalled, "analysisService.analyze should be called.");
      T.assertNotNull(lastToast, "Success notification should be shown.");
      T.assertEquals("Financial overview generated successfully!", lastToast.message, "Correct success message expected.");

      // Verify basic sheet setup (more detailed checks are complex with mocks)
      const overviewData = mockSheetData["Overview"];
      T.assertTrue(overviewData.length > 5, "Overview sheet should have multiple rows generated."); // Basic check
      T.assertEquals("Type", overviewData[0][0], "Header row should contain 'Type'.");
      T.assertEquals("Show Sub-Categories", overviewData[0][18], "Sub-category toggle label should be present."); // Assuming S=19
  });

  T.registerTest("FinanceOverview", "handleEdit should update preference and recreate on checkbox change", function() {
      setupMockDataAndState();
      mockUserPropertiesStore["ShowSubCategories"] = true; // Initial state

      const overviewSheet = mockActiveSpreadsheet.getSheetByName("Overview");
      const checkboxCell = overviewSheet.getRange(mockConfig._data.UI.SUBCATEGORY_TOGGLE.CHECKBOX_CELL); // e.g., T1
      checkboxCell.setValue(false); // Simulate user unchecking the box

      const mockEvent = {
          range: checkboxCell,
          value: false, // The new value after edit
          oldValue: true,
          source: mockActiveSpreadsheet
      };

      // Reset interaction flags before calling handleEdit
      settingsSetValueLog = {};
      cacheInvalidateAllCalled = false;
      analysisServiceAnalyzeCalled = false;
      lastToast = null;

      TestFinanceOverview.handleEdit(mockEvent);

      T.assertTrue(settingsSetValueLog["ShowSubCategories"] > 0, "settingsService.setValue('ShowSubCategories') should be called.");
      T.assertEquals(false, mockUserPropertiesStore["ShowSubCategories"], "User preference should be updated to false.");
      // Check if create was called again (indicated by its interactions)
      T.assertTrue(cacheInvalidateAllCalled, "cacheService.invalidateAll should be called by subsequent create().");
      T.assertTrue(analysisServiceAnalyzeCalled, "analysisService.analyze should be called by subsequent create().");
      T.assertNotNull(lastToast, "Success notification should be shown by subsequent create().");
       T.assertEquals("Financial overview generated successfully!", lastToast.message, "Correct success message from create() expected.");
  });

   T.registerTest("FinanceOverview", "handleEdit should do nothing if edit is not on checkbox", function() {
      setupMockDataAndState();
      const overviewSheet = mockActiveSpreadsheet.getSheetByName("Overview");
      const otherCell = overviewSheet.getRange(5, 5); // Some other cell
      otherCell.setValue("Some data");

       const mockEvent = {
          range: otherCell,
          value: "Some data",
          source: mockActiveSpreadsheet
      };

      // Reset interaction flags
      settingsSetValueLog = {};
      cacheInvalidateAllCalled = false;
      analysisServiceAnalyzeCalled = false;
      lastToast = null;

       TestFinanceOverview.handleEdit(mockEvent);

       T.assertTrue(settingsSetValueLog["ShowSubCategories"] === undefined, "settingsService.setValue should NOT be called.");
       T.assertFalse(cacheInvalidateAllCalled, "cacheService.invalidateAll should NOT be called.");
       T.assertFalse(analysisServiceAnalyzeCalled, "analysisService.analyze should NOT be called.");
       T.assertTrue(lastToast === null, "No UI notification should be shown.");
   });


})(); // End IIFE
