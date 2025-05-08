/**
 * Financial Planning Tools - Dropdown Service Tests
 *
 * This file contains tests for the FinancialPlanner.DropdownService module.
 * It includes mocking SpreadsheetApp, CacheService, and other dependencies.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;

  // --- Mock Dependencies & Globals ---
  let mockSheetData = {}; // { sheetName: [[row1], [row2], ...] }
  let mockActiveSpreadsheet = null;
  let mockScriptCacheStore = {};
  let lastToast = null;
  let lastHandledError = null;
  let dataValidations = {}; // { sheetName: { rangeA1: rule } }

  // --- Mock Spreadsheet Objects ---
   const mockRange = {
    _sheetName: null, _row: 0, _col: 0, _numRows: 1, _numCols: 1, _sheetDataRef: null,
    _validationRule: null, _background: null,

    setValue: function(value) { /* ... simplified ... */
        if (this._numRows === 1 && this._numCols === 1) this._sheetDataRef[this._row - 1][this._col - 1] = value;
        else throw new Error("MockRange: setValue only for single cells.");
        return this;
    },
    getValue: function() { /* ... simplified ... */
        if (this._numRows === 1 && this._numCols === 1) return this._sheetDataRef[this._row - 1][this._col - 1];
        throw new Error("MockRange: getValue only for single cells.");
    },
     getValues: function() {
        const result = [];
        for(let r=0; r<this._numRows; r++) {
            const rowData = [];
            // Ensure row exists
            if (!this._sheetDataRef[this._row + r - 1]) this._sheetDataRef[this._row + r - 1] = [];
            for (let c=0; c<this._numCols; c++) {
                 // Ensure cell exists
                rowData.push(this._sheetDataRef[this._row + r - 1][this._col + c - 1] || "");
            }
            result.push(rowData);
        }
        return result;
    },
    clearContent: function() { /* ... simplified ... */
        if (this._numRows === 1 && this._numCols === 1) this._sheetDataRef[this._row - 1][this._col - 1] = "";
        else { // Clear range
             for(let r=0; r<this._numRows; r++) {
                for (let c=0; c<this._numCols; c++) {
                     if (this._sheetDataRef[this._row + r - 1]) {
                         this._sheetDataRef[this._row + r - 1][this._col + c - 1] = "";
                     }
                }
             }
        }
        return this;
    },
    clearDataValidations: function() {
        const rangeA1 = `R${this._row}C${this._col}`; // Simple key for mock
        if(dataValidations[this._sheetName]) delete dataValidations[this._sheetName][rangeA1];
        this._validationRule = null;
        return this;
    },
    setDataValidation: function(rule) {
        const rangeA1 = `R${this._row}C${this._col}`; // Simple key for mock
        if(!dataValidations[this._sheetName]) dataValidations[this._sheetName] = {};
        dataValidations[this._sheetName][rangeA1] = rule; // Store the mock rule
        this._validationRule = rule;
        return this;
    },
    setBackground: function(color) { this._background = color; return this; },
    getA1Notation: function() { return `R${this._row}C${this._col}`; }, // Simplified A1 for tests
    getSheet: function() { return mockActiveSpreadsheet.getSheetByName(this._sheetName); } // Return mock sheet
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
    getFrozenRows: function() { return this._frozenRows; }
    // Add other methods if needed
  };

  const mockDataValidation = { _values: [], _allowInvalid: true };
  const mockDataValidationBuilder = {
      requireValueInList: function(values, showDropdown) { mockDataValidation._values = values; return this; },
      setAllowInvalid: function(allow) { mockDataValidation._allowInvalid = allow; return this; },
      build: function() { return Object.assign({}, mockDataValidation); } // Return a copy
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
    newDataValidation: function() { return Object.assign({}, mockDataValidationBuilder); } // Return new builder instance
  };
  global.CacheService = {
    getScriptCache: function() {
        return {
            get: function(key){ return mockScriptCacheStore[key] || null; },
            put: function(key, value, ttl){ mockScriptCacheStore[key] = value; }, // Simplified put
            remove: function(key){ delete mockScriptCacheStore[key]; }
        };
    }
  };
  global.Logger = { log: function(msg) { console.log("Logger.log:", msg); } }; // Basic Logger mock

  // --- Mock Other Dependencies ---
  const mockConfig = {
    _sheets: { TRANSACTIONS: "Transactions", DROPDOWNS: "Dropdowns" },
    get: function() { return { SHEETS: this._sheets }; }, // Simplified get
    getSection: function(section) { if (section === 'SHEETS') return this._sheets; return {}; }
  };
  const mockUtils = { /* Mock specific utils if needed, assume none for now */ };
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
   // Redefine DropdownService with mocks
   const TestDropdownService = (function(utils, uiService, errorService, config) {
       // --- Copy of DropdownService Implementation Start ---
        const DROPDOWN_CONFIG = { CACHE_EXPIRY_SECONDS: 300, CACHE_KEY: 'dropdownsData', SHEETS: { TRANSACTIONS: 'Transactions', DROPDOWNS: 'Dropdowns' }, COLUMNS: { TYPE: 3, CATEGORY: 4, SUB_CATEGORY: 5 }, UI: { PLACEHOLDER_TEXT: 'Please select', PENDING_BACKGROUND: '#eeeeee' }, KEY_SEPARATOR: '___' };
        let dropdownCache = null;
        function mapSetsToArrays(obj) { const out = {}; for (let key in obj) { if (obj.hasOwnProperty(key)) { out[key] = Array.from(obj[key]); } } return out; }
        function buildDropdownCache(spreadsheet) { try { const cache = CacheService.getScriptCache(); let cachedData = cache.get(DROPDOWN_CONFIG.CACHE_KEY); if (cachedData) { Logger.log('Using cached dropdown data'); const parsed = JSON.parse(cachedData); return { typeToCategories: parsed.typeToCategories || {}, typeCategoryToSubCategories: parsed.typeCategoryToSubCategories || {}, }; } Logger.log('Building dropdown cache from sheet'); const dropdownsSheetName = DROPDOWN_CONFIG.SHEETS.DROPDOWNS; const dropdownsSheet = spreadsheet.getSheetByName(dropdownsSheetName); if (!dropdownsSheet) { throw new Error(`Sheet "${dropdownsSheetName}" not found`); } const dropdownsData = dropdownsSheet.getDataRange().getValues(); const startRow = (dropdownsData.length > 0 && typeof dropdownsData[0][0] === 'string' && dropdownsData[0][0].toLowerCase() === 'type') ? 1 : 0; const typeToCategories = {}; const typeCategoryToSubCategories = {}; for (let i = startRow; i < dropdownsData.length; i++) { const [type, category, subCategory] = dropdownsData[i]; if (!type || !category) continue; if (!typeToCategories[type]) { typeToCategories[type] = new Set(); } typeToCategories[type].add(category); if (subCategory) { const key = `${type}${DROPDOWN_CONFIG.KEY_SEPARATOR}${category}`; if (!typeCategoryToSubCategories[key]) { typeCategoryToSubCategories[key] = new Set(); } typeCategoryToSubCategories[key].add(subCategory); } } const cacheToStore = { typeToCategories: mapSetsToArrays(typeToCategories), typeCategoryToSubCategories: mapSetsToArrays(typeCategoryToSubCategories), }; try { cache.put(DROPDOWN_CONFIG.CACHE_KEY, JSON.stringify(cacheToStore), DROPDOWN_CONFIG.CACHE_EXPIRY_SECONDS); } catch (e) { Logger.log('Failed to cache dropdown data: ' + e.toString()); } return cacheToStore; } catch (error) { Logger.log('Error building dropdown cache: ' + error.toString()); errorService.log(errorService.create('Error building dropdown cache', { originalError: error.toString() })); return { typeToCategories: {}, typeCategoryToSubCategories: {} }; } }
        function setDropdownValidation(range, options) { range.clearDataValidations(); if (!options || options.length === 0) { return; } const rule = SpreadsheetApp.newDataValidation().requireValueInList([DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT, ...options], true).setAllowInvalid(true).build(); range.setDataValidation(rule); }
        function clearPlaceholderIfNeeded(range) { if (range.getValue() === DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT) { range.clearContent(); } }
        function highlightPending(range) { range.setBackground(DROPDOWN_CONFIG.UI.PENDING_BACKGROUND); }
        function clearHighlight(range) { range.setBackground(null); }
        function handleTypeEdit(sheet, row, typeToCategories) { const typeCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.TYPE); const categoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.CATEGORY); const subCategoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.SUB_CATEGORY); const type = typeCell.getValue(); const oldCategoryValue = categoryCell.getValue(); const categories = typeToCategories[type] || []; setDropdownValidation(categoryCell, categories); if (oldCategoryValue && (categories.length === 0 || !categories.includes(oldCategoryValue))) { categoryCell.clearContent(); subCategoryCell.clearContent().clearDataValidations(); highlightPending(categoryCell); } else if (categories.length > 0 && !oldCategoryValue) { highlightPending(categoryCell); } clearPlaceholderIfNeeded(typeCell); if (categoryCell.getValue() !== DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT && categoryCell.getValue() !== "") { clearHighlight(categoryCell); } }
        function handleCategoryEdit(sheet, row, typeCategoryToSubCategories) { const typeCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.TYPE); const categoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.CATEGORY); const subCategoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.SUB_CATEGORY); const type = typeCell.getValue(); const category = categoryCell.getValue(); const oldSubCategoryValue = subCategoryCell.getValue(); if (type && category && category !== DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT) { const key = `${type}${DROPDOWN_CONFIG.KEY_SEPARATOR}${category}`; const subCategories = typeCategoryToSubCategories[key] || []; setDropdownValidation(subCategoryCell, subCategories); if (oldSubCategoryValue && (subCategories.length === 0 || !subCategories.includes(oldSubCategoryValue))) { subCategoryCell.clearContent(); highlightPending(subCategoryCell); } else if (subCategories.length > 0 && !oldSubCategoryValue) { highlightPending(subCategoryCell); } } else { subCategoryCell.clearContent().clearDataValidations(); } clearPlaceholderIfNeeded(categoryCell); if (subCategoryCell.getValue() !== DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT && subCategoryCell.getValue() !== "") { clearHighlight(subCategoryCell); } }
        function handleSubCategoryEdit(sheet, row) { const subCategoryCell = sheet.getRange(row, DROPDOWN_CONFIG.COLUMNS.SUB_CATEGORY); clearPlaceholderIfNeeded(subCategoryCell); clearHighlight(subCategoryCell); }
        return { handleEdit: function(e) { try { const sheet = e.range.getSheet(); const transactionsSheetName = (config && config.get && config.get().SHEETS && config.get().SHEETS.TRANSACTIONS) ? config.get().SHEETS.TRANSACTIONS : DROPDOWN_CONFIG.SHEETS.TRANSACTIONS; if (sheet.getName() !== transactionsSheetName) return; if (e.range.getRow() === 1 && sheet.getFrozenRows() >= 1) return; if (!dropdownCache) { dropdownCache = buildDropdownCache(e.source); } const { typeToCategories, typeCategoryToSubCategories } = dropdownCache; const startRow = e.range.getRow(); const startCol = e.range.getColumn(); const numRows = e.range.getNumRows(); const numCols = e.range.getNumColumns(); for (let r = 0; r < numRows; r++) { const currentRow = startRow + r; if (currentRow === 1 && sheet.getFrozenRows() >= 1) continue; for (let c = 0; c < numCols; c++) { const currentCol = startCol + c; if (currentCol === DROPDOWN_CONFIG.COLUMNS.TYPE) { handleTypeEdit(sheet, currentRow, typeToCategories); } else if (currentCol === DROPDOWN_CONFIG.COLUMNS.CATEGORY) { handleCategoryEdit(sheet, currentRow, typeCategoryToSubCategories); } else if (currentCol === DROPDOWN_CONFIG.COLUMNS.SUB_CATEGORY) { handleSubCategoryEdit(sheet, currentRow); } } } } catch (error) { Logger.log('Error in DropdownService.handleEdit: ' + error.toString()); errorService.handle(errorService.create('Error in dropdown edit handler', { originalError: error.toString() }), "An error occurred while updating dropdowns."); } }, refreshCache: function() { try { uiService.showLoadingSpinner("Refreshing dropdown cache..."); const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); const scriptCache = CacheService.getScriptCache(); scriptCache.remove(DROPDOWN_CONFIG.CACHE_KEY); dropdownCache = buildDropdownCache(spreadsheet); uiService.hideLoadingSpinner(); uiService.showSuccessNotification('Dropdown cache has been refreshed.'); } catch (error) { uiService.hideLoadingSpinner(); Logger.log('Error refreshing dropdown cache: ' + error.toString()); errorService.handle(errorService.create('Error refreshing dropdown cache', { originalError: error.toString() }), "Failed to refresh dropdown cache."); } }, initializeCache: function() { if (!dropdownCache) { Logger.log("Initializing dropdown cache proactively."); dropdownCache = buildDropdownCache(SpreadsheetApp.getActiveSpreadsheet()); } } };
       // --- Copy of DropdownService Implementation End ---
   })(mockUtils, mockUiService, mockErrorService, mockConfig); // Pass mocks

  // --- Helper to reset state before each test ---
  function setupMockSheets() {
      mockSheetData = { // Reset sheet data
          "Dropdowns": [
              ["Type", "Category", "Sub-Category"],
              ["Expense", "Food", "Groceries"],
              ["Expense", "Food", "Restaurants"],
              ["Expense", "Utilities", "Electricity"],
              ["Expense", "Utilities", "Water"],
              ["Income", "Salary", ""],
              ["Income", "Bonus", ""]
          ],
          "Transactions": [
              ["Date", "Description", "Type", "Category", "Sub-Category", "Amount"],
              ["2024-01-10", "Grocery Store", "Expense", "Food", "Groceries", 50],
              ["2024-01-12", "Restaurant Bill", "Expense", "Food", "Restaurants", 75],
              ["2024-01-15", "Paycheck", "Income", "Salary", "", 2000]
          ]
      };
      mockActiveSpreadsheet.sheets = {}; // Clear existing sheets
      mockActiveSpreadsheet.insertSheet("Dropdowns");
      mockActiveSpreadsheet.insertSheet("Transactions");
      mockScriptCacheStore = {}; // Clear script cache
      dataValidations = {}; // Clear validations
      lastToast = null;
      lastHandledError = null;
      // Reset internal cache of the service
      TestDropdownService.refreshCache(); // Use refresh to clear and rebuild internal cache
      lastToast = null; // Clear toast from refreshCache
  }

  // --- Test Cases ---

  T.registerTest("DropdownService", "initializeCache should build cache if empty", function() {
      setupMockSheets();
      // Ensure internal cache is null initially (refreshCache sets it, so we manually nullify for this test)
      TestDropdownService.dropdownCache = null; // Accessing internal state for test setup
      mockScriptCacheStore = {}; // Ensure script cache is also empty

      TestDropdownService.initializeCache();
      // Check internal cache (indirectly, by seeing if handleEdit works without rebuild)
      const mockEvent = {
          range: mockActiveSpreadsheet.getSheetByName("Transactions").getRange(2, 3), // Edit Type in R2
          source: mockActiveSpreadsheet
      };
      // If initializeCache worked, handleEdit shouldn't call buildDropdownCache again (check logs or add counter if needed)
      TestDropdownService.handleEdit(mockEvent);
      // Basic check: did it run without error?
      T.assertTrue(lastHandledError === null, "initializeCache followed by handleEdit should not cause errors.");
  });

  T.registerTest("DropdownService", "refreshCache should clear script cache and rebuild internal cache", function() {
      setupMockSheets();
      // Populate script cache
      mockScriptCacheStore[TestDropdownService.DROPDOWN_CONFIG.CACHE_KEY] = JSON.stringify({ typeToCategories: { "Test": ["Cached"] } });
      // Initialize internal cache (might already be done by setup)
      TestDropdownService.initializeCache();

      TestDropdownService.refreshCache();

      // Check script cache was cleared (or attempted to be cleared via remove)
      T.assertTrue(mockScriptCacheStore[TestDropdownService.DROPDOWN_CONFIG.CACHE_KEY] === undefined, "Script cache should be cleared by refreshCache.");
      // Check internal cache was rebuilt (verify content based on mock sheet data)
      const internalCache = TestDropdownService.dropdownCache; // Access internal state for test verification
      T.assertNotNull(internalCache, "Internal cache should be rebuilt.");
      T.assertTrue(!!internalCache.typeToCategories["Expense"], "Internal cache should contain 'Expense' type after rebuild.");
      T.assertTrue(internalCache.typeToCategories["Expense"].includes("Food"), "Expense categories should include 'Food'.");
      T.assertNotNull(lastToast, "Success notification should be shown after refresh.");
      T.assertEquals("Dropdown cache has been refreshed.", lastToast.message, "Correct success message expected.");
  });

  T.registerTest("DropdownService", "handleEdit on Type should update Category dropdown", function() {
      setupMockSheets();
      const transactionsSheet = mockActiveSpreadsheet.getSheetByName("Transactions");
      const editRange = transactionsSheet.getRange(4, 3); // New row (R4), Type column (C3)
      editRange.setValue("Expense"); // Set the type value that triggers the edit

      const mockEvent = { range: editRange, source: mockActiveSpreadsheet, value: "Expense" };
      TestDropdownService.handleEdit(mockEvent);

      // Check validation on Category cell (R4, C4)
      const categoryCellA1 = `R4C4`;
      const validationRule = dataValidations["Transactions"] ? dataValidations["Transactions"][categoryCellA1] : null;
      T.assertNotNull(validationRule, "Data validation should be set on Category cell.");
      // Check options (includes placeholder + sheet data)
      T.assertTrue(validationRule._values.includes("Food"), "Category options should include 'Food'.");
      T.assertTrue(validationRule._values.includes("Utilities"), "Category options should include 'Utilities'.");
      T.assertEquals(TestDropdownService.DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT, validationRule._values[0], "First option should be placeholder.");
      // Check highlighting
      T.assertEquals(TestDropdownService.DROPDOWN_CONFIG.UI.PENDING_BACKGROUND, transactionsSheet.getRange(4, 4)._background, "Category cell should be highlighted.");
  });

   T.registerTest("DropdownService", "handleEdit on Category should update Sub-Category dropdown", function() {
      setupMockSheets();
      const transactionsSheet = mockActiveSpreadsheet.getSheetByName("Transactions");
      // Pre-set Type in R4
      transactionsSheet.getRange(4, 3).setValue("Expense");
      // Edit Category in R4
      const editRange = transactionsSheet.getRange(4, 4);
      editRange.setValue("Food");

      const mockEvent = { range: editRange, source: mockActiveSpreadsheet, value: "Food" };
      TestDropdownService.handleEdit(mockEvent);

      // Check validation on Sub-Category cell (R4, C5)
      const subCategoryCellA1 = `R4C5`;
      const validationRule = dataValidations["Transactions"] ? dataValidations["Transactions"][subCategoryCellA1] : null;
      T.assertNotNull(validationRule, "Data validation should be set on Sub-Category cell.");
      T.assertTrue(validationRule._values.includes("Groceries"), "Sub-Category options should include 'Groceries'.");
      T.assertTrue(validationRule._values.includes("Restaurants"), "Sub-Category options should include 'Restaurants'.");
      T.assertEquals(TestDropdownService.DROPDOWN_CONFIG.UI.PLACEHOLDER_TEXT, validationRule._values[0], "First sub-category option should be placeholder.");
       // Check highlighting
      T.assertEquals(TestDropdownService.DROPDOWN_CONFIG.UI.PENDING_BACKGROUND, transactionsSheet.getRange(4, 5)._background, "Sub-Category cell should be highlighted.");
   });
   
    T.registerTest("DropdownService", "handleEdit on SubCategory should clear highlight", function() {
      setupMockSheets();
      const transactionsSheet = mockActiveSpreadsheet.getSheetByName("Transactions");
      // Pre-set Type and Category, and highlight
      transactionsSheet.getRange(4, 3).setValue("Expense");
      transactionsSheet.getRange(4, 4).setValue("Food");
      transactionsSheet.getRange(4, 5).setBackground(TestDropdownService.DROPDOWN_CONFIG.UI.PENDING_BACKGROUND); // Simulate highlight

      // Edit Sub-Category in R4
      const editRange = transactionsSheet.getRange(4, 5);
      editRange.setValue("Groceries");

      const mockEvent = { range: editRange, source: mockActiveSpreadsheet, value: "Groceries" };
      TestDropdownService.handleEdit(mockEvent);

      // Check highlight on Sub-Category cell (R4, C5)
      T.assertEquals(null, transactionsSheet.getRange(4, 5)._background, "Sub-Category cell highlight should be cleared.");
   });
   
    T.registerTest("DropdownService", "handleEdit changing Type should clear invalid Category/SubCategory", function() {
      setupMockSheets();
      const transactionsSheet = mockActiveSpreadsheet.getSheetByName("Transactions");
      // Pre-set row 2: Expense, Food, Groceries
      transactionsSheet.getRange(2, 3).setValue("Expense");
      transactionsSheet.getRange(2, 4).setValue("Food");
      transactionsSheet.getRange(2, 5).setValue("Groceries");
       // Set validation for sub-category based on initial state
       const initialSubCatRule = SpreadsheetApp.newDataValidation().requireValueInList(['Please select', 'Groceries', 'Restaurants'], true).build();
       transactionsSheet.getRange(2, 5).setDataValidation(initialSubCatRule);


      // Edit Type in R2 to Income
      const editRange = transactionsSheet.getRange(2, 3);
      editRange.setValue("Income");

      const mockEvent = { range: editRange, source: mockActiveSpreadsheet, value: "Income" };
      TestDropdownService.handleEdit(mockEvent);

      // Check Category (R2, C4) - should be cleared and highlighted
      T.assertEquals("", transactionsSheet.getRange(2, 4).getValue(), "Category should be cleared as 'Food' is not valid for 'Income'.");
      T.assertEquals(TestDropdownService.DROPDOWN_CONFIG.UI.PENDING_BACKGROUND, transactionsSheet.getRange(2, 4)._background, "Category cell should be highlighted after Type change makes it invalid.");
      // Check Sub-Category (R2, C5) - should be cleared and validation removed
      T.assertEquals("", transactionsSheet.getRange(2, 5).getValue(), "Sub-Category should be cleared.");
       const subCategoryCellA1 = `R2C5`;
       const validationRule = dataValidations["Transactions"] ? dataValidations["Transactions"][subCategoryCellA1] : null;
       T.assertTrue(validationRule === null || validationRule._values.length <= 1, "Sub-Category validation should be cleared or only contain placeholder."); // Check validation cleared
   });


})(); // End IIFE
