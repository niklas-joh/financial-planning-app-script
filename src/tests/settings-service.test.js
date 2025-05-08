/**
 * Financial Planning Tools - Settings Service Tests
 *
 * This file contains tests for the FinancialPlanner.SettingsService module.
 * It includes mocking SpreadsheetApp interactions.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;

  // --- Mock Dependencies & Globals ---

  // Mock SpreadsheetApp and related classes
  const mockSheetData = {}; // Store mock sheet data { sheetName: [[row1], [row2], ...] }
  let mockActiveSpreadsheet = null; // Reference to the mock spreadsheet

  const mockRange = {
    _sheetName: null,
    _row: 0,
    _col: 0,
    _numRows: 1,
    _numCols: 1,
    _value: null,
    _values: null,
    _sheetDataRef: null, // Reference to the specific sheet's data in mockSheetData

    setValue: function(value) {
      if (this._numRows === 1 && this._numCols === 1) {
        this._sheetDataRef[this._row - 1][this._col - 1] = value;
      } else {
         // Simplified: only handle single cell setValue for these tests
         throw new Error("MockRange: setValue only implemented for single cells in this mock.");
      }
      return this;
    },
    setValues: function(values) {
       if (values.length !== this._numRows || values[0].length !== this._numCols) {
           throw new Error("MockRange: setValues dimensions mismatch.");
       }
       for(let r=0; r<this._numRows; r++) {
           for (let c=0; c<this._numCols; c++) {
               this._sheetDataRef[this._row + r - 1][this._col + c - 1] = values[r][c];
           }
       }
       return this;
    },
    getValue: function() {
       if (this._numRows === 1 && this._numCols === 1) {
           return this._sheetDataRef[this._row - 1][this._col - 1];
       }
       throw new Error("MockRange: getValue only implemented for single cells.");
    },
    getValues: function() {
        const result = [];
        for(let r=0; r<this._numRows; r++) {
            const rowData = [];
            for (let c=0; c<this._numCols; c++) {
                rowData.push(this._sheetDataRef[this._row + r - 1][this._col + c - 1]);
            }
            result.push(rowData);
        }
        return result;
    },
    setFontWeight: function() { return this; }, // Chainable mocks
    clearContent: function() {
        for(let r=0; r<this._numRows; r++) {
           for (let c=0; c<this._numCols; c++) {
               // Avoid clearing headers if range starts at row 1
               if (this._row + r > 1) {
                  this._sheetDataRef[this._row + r - 1][this._col + c - 1] = "";
               }
           }
       }
       return this;
    }
  };

  const mockSheet = {
    _name: null,
    _dataRef: null, // Reference to this sheet's data in mockSheetData
    _hidden: false,

    getName: function() { return this._name; },
    getDataRange: function() {
      // Simplified: return range covering all data
      const numRows = this._dataRef.length;
      const numCols = numRows > 0 ? this._dataRef[0].length : 0;
      return Object.assign({}, mockRange, {
          _sheetName: this._name, _row: 1, _col: 1, _numRows: numRows, _numCols: numCols, _sheetDataRef: this._dataRef
      });
    },
    getRange: function(row, col, numRows = 1, numCols = 1) {
       // Ensure sheet has enough rows/cols - simplified mock doesn't auto-expand
       while(this._dataRef.length < row + numRows -1) this._dataRef.push([]);
       const maxCols = col + numCols -1;
       this._dataRef.forEach(r => { while(r.length < maxCols) r.push(""); });

       return Object.assign({}, mockRange, {
           _sheetName: this._name, _row: row, _col: col, _numRows: numRows, _numCols: numCols, _sheetDataRef: this._dataRef
       });
    },
    appendRow: function(rowData) {
        this._dataRef.push([...rowData]); // Add a copy
        // Ensure consistent column count
        const maxCols = this._dataRef[0] ? this._dataRef[0].length : rowData.length;
         this._dataRef.forEach(r => { while(r.length < maxCols) r.push(""); });
    },
    insertSheet: function(sheetName) { return mockActiveSpreadsheet.insertSheet(sheetName); }, // Delegate
    hideSheet: function() { this._hidden = true; },
    getLastRow: function() { return this._dataRef.length; },
    getMaxColumns: function() { return this._dataRef[0] ? this._dataRef[0].length : 0; }
  };

  mockActiveSpreadsheet = {
    sheets: {}, // Store mock sheets by name
    getSheetByName: function(name) {
      return this.sheets[name] || null;
    },
    insertSheet: function(name) {
      if (this.sheets[name]) return this.sheets[name]; // Already exists
      const newSheetData = [];
      mockSheetData[name] = newSheetData;
      const newSheet = Object.assign({}, mockSheet, { _name: name, _dataRef: newSheetData });
      this.sheets[name] = newSheet;
      return newSheet;
    },
    getActiveSpreadsheet: function() { return this; } // Return self
  };

  // Global mock
  global.SpreadsheetApp = {
    getActiveSpreadsheet: function() { return mockActiveSpreadsheet; },
    // Add other SpreadsheetApp methods if needed by the service
  };
  
  // --- Mock Other Dependencies ---
  const mockConfig = {
    _settingsSheetName: "Test Settings",
    getSheetNames: function() { return { SETTINGS: this._settingsSheetName }; },
     getSection: function(section) { // Added for compatibility
        if (section === 'SHEETS') return this.getSheetNames();
        return {};
    }
  };

  const mockUtils = {
    // Use the actual getOrCreateSheet logic but with mocked SpreadsheetApp
     getOrCreateSheet: function(spreadsheet, sheetName) {
      let sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
      } else {
        // Don't clear content in tests unless specifically testing reset
        // sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
      }
      return sheet;
    }
  };

  let lastSuccessNotification = null;
  const mockUiService = {
    showSuccessNotification: function(message) { lastSuccessNotification = message; },
    // Add other methods if needed
  };

  let lastHandledError = null;
  const mockErrorService = {
    handle: function(error, message) { lastHandledError = { error: error, message: message }; },
    create: function(message, details) { // Simple error object creation
        const err = new Error(message);
        err.details = details;
        err.name = 'FinancialPlannerError'; // Simulate custom error
        return err;
     }
  };

  // --- Test Suite Setup ---
  // Redefine the service for testing, injecting mocks
   const TestSettingsService = (function(config, utils, uiService, errorService) {
      // --- Copy of SettingsService Implementation Start ---
        function getSettingsSheet() {
            const ss = SpreadsheetApp.getActiveSpreadsheet(); // Uses mock
            const sheetName = config.getSheetNames().SETTINGS;
            let sheet = ss.getSheetByName(sheetName);
            if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            sheet.getRange("A1:B1").setValues([["Preference", "Value"]]).setFontWeight("bold");
            sheet.hideSheet();
            }
            return sheet;
        }

        function findPreference(key) {
            const sheet = getSettingsSheet();
            const data = sheet.getDataRange().getValues();
            for (let i = 1; i < data.length; i++) {
            if (data[i][0] === key) {
                return { row: i + 1, value: data[i][1] };
            }
            }
            return null;
        }

        return {
            getValue: function(key, defaultValue) {
                try {
                    const preference = findPreference(key);
                    return preference ? preference.value : defaultValue;
                } catch (error) {
                    errorService.handle(errorService.create(`Error getting setting value for key: ${key}`, { originalError: error.toString() }), `Failed to get setting: ${key}`);
                    return defaultValue;
                }
            },
            setValue: function(key, value) {
                try {
                    const sheet = getSettingsSheet();
                    const preference = findPreference(key);
                    if (preference) {
                    sheet.getRange(preference.row, 2).setValue(value);
                    } else {
                    const lastRow = Math.max(1, sheet.getLastRow());
                    sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
                    }
                } catch (error) {
                    errorService.handle(errorService.create(`Error setting setting value for key: ${key}`, { originalError: error.toString(), valueToSet: value }), `Failed to set setting: ${key}`);
                }
            },
             getBooleanValue: function(key, defaultValue) {
                const value = this.getValue(key, defaultValue);
                if (typeof value === 'boolean') return value;
                if (value === 'true' || value === 1 || value === '1') return true;
                if (value === 'false' || value === 0 || value === '0') return false;
                return !!defaultValue;
            },
            getNumericValue: function(key, defaultValue) {
                const value = this.getValue(key, defaultValue);
                if (typeof value === 'number') return value;
                const parsed = parseFloat(value);
                return isNaN(parsed) ? (typeof defaultValue === 'number' ? defaultValue : 0) : parsed;
            },
            toggleBooleanValue: function(key, defaultValue) {
                const currentValue = this.getBooleanValue(key, defaultValue);
                const newValue = !currentValue;
                this.setValue(key, newValue);
                return newValue;
            },
             toggleShowSubCategories: function() {
                return this.toggleBooleanValue("ShowSubCategories", true);
            },
            getShowSubCategories: function() {
                return this.getBooleanValue("ShowSubCategories", true);
            },
            setShowSubCategories: function(value) {
                this.setValue("ShowSubCategories", typeof value === 'boolean' ? value : true);
            },
            getAllPreferences: function() {
                try {
                    const sheet = getSettingsSheet();
                    const data = sheet.getDataRange().getValues();
                    const preferences = {};
                    for (let i = 1; i < data.length; i++) {
                    if (data[i] && data[i][0] != null && data[i][0] !== "") {
                        preferences[data[i][0]] = data[i][1];
                    }
                    }
                    return preferences;
                } catch (error) {
                    errorService.handle(errorService.create('Error getting all preferences', { originalError: error.toString() }), "Failed to retrieve all settings.");
                    return {};
                }
            },
            resetAllPreferences: function() {
                try {
                    const sheet = getSettingsSheet();
                    if (sheet.getLastRow() > 1) {
                    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
                    }
                    uiService.showSuccessNotification("All preferences have been reset.");
                } catch (error) {
                    errorService.handle(errorService.create('Error resetting all preferences', { originalError: error.toString() }), "Failed to reset settings.");
                }
            }
        };
      // --- Copy of SettingsService Implementation End ---
   })(mockConfig, mockUtils, mockUiService, mockErrorService); // Pass mocks


  // --- Helper to reset mock sheet state before each test ---
  function setupMockSheet() {
      mockSheetData[mockConfig._settingsSheetName] = [["Preference", "Value"]]; // Reset with header
      mockActiveSpreadsheet.sheets = {}; // Clear existing sheets
      mockActiveSpreadsheet.insertSheet(mockConfig._settingsSheetName); // Re-create the settings sheet
      lastSuccessNotification = null;
      lastHandledError = null;
  }

  // --- Test Cases ---

  T.registerTest("SettingsService", "getValue should return default value if key not found", function() {
    setupMockSheet();
    const defaultValue = "default";
    const value = TestSettingsService.getValue("nonexistent_key", defaultValue);
    T.assertEquals(defaultValue, value, "Should return default value for nonexistent key.");
  });

  T.registerTest("SettingsService", "setValue should add a new preference", function() {
    setupMockSheet();
    const key = "new_pref";
    const value = "new_value";
    TestSettingsService.setValue(key, value);
    const retrievedValue = TestSettingsService.getValue(key, "default");
    T.assertEquals(value, retrievedValue, "Should retrieve the newly set value.");
    // Check sheet data directly (simplified)
    const sheetData = mockSheetData[mockConfig._settingsSheetName];
    T.assertEquals(2, sheetData.length, "Sheet should have header + 1 data row.");
    T.assertEquals(key, sheetData[1][0], "Key should be in the second row, first column.");
    T.assertEquals(value, sheetData[1][1], "Value should be in the second row, second column.");
  });

  T.registerTest("SettingsService", "setValue should update an existing preference", function() {
    setupMockSheet();
    const key = "existing_pref";
    const initialValue = "initial_value";
    const updatedValue = "updated_value";
    // Set initial value
    mockSheetData[mockConfig._settingsSheetName].push([key, initialValue]);

    TestSettingsService.setValue(key, updatedValue);
    const retrievedValue = TestSettingsService.getValue(key, "default");
    T.assertEquals(updatedValue, retrievedValue, "Should retrieve the updated value.");
     // Check sheet data directly
    const sheetData = mockSheetData[mockConfig._settingsSheetName];
    T.assertEquals(2, sheetData.length, "Sheet should still have header + 1 data row.");
    T.assertEquals(key, sheetData[1][0], "Key should be in the second row, first column.");
    T.assertEquals(updatedValue, sheetData[1][1], "Value should be updated in the second row, second column.");
  });

  T.registerTest("SettingsService", "getBooleanValue should coerce values correctly", function() {
    setupMockSheet();
    TestSettingsService.setValue("bool_true", true);
    TestSettingsService.setValue("bool_false", false);
    TestSettingsService.setValue("string_true", "true");
    TestSettingsService.setValue("string_false", "false");
    TestSettingsService.setValue("num_1", 1);
    TestSettingsService.setValue("num_0", 0);
    TestSettingsService.setValue("string_1", "1");
    TestSettingsService.setValue("string_0", "0");
    TestSettingsService.setValue("other_string", "hello");

    T.assertTrue(TestSettingsService.getBooleanValue("bool_true", false), "Boolean true should return true.");
    T.assertFalse(TestSettingsService.getBooleanValue("bool_false", true), "Boolean false should return false.");
    T.assertTrue(TestSettingsService.getBooleanValue("string_true", false), "String 'true' should return true.");
    T.assertFalse(TestSettingsService.getBooleanValue("string_false", true), "String 'false' should return false.");
    T.assertTrue(TestSettingsService.getBooleanValue("num_1", false), "Number 1 should return true.");
    T.assertFalse(TestSettingsService.getBooleanValue("num_0", true), "Number 0 should return false.");
     T.assertTrue(TestSettingsService.getBooleanValue("string_1", false), "String '1' should return true.");
    T.assertFalse(TestSettingsService.getBooleanValue("string_0", true), "String '0' should return false.");
    T.assertFalse(TestSettingsService.getBooleanValue("other_string", false), "Other string should return default (false).");
    T.assertTrue(TestSettingsService.getBooleanValue("nonexistent", true), "Nonexistent key should return default (true).");
  });

  T.registerTest("SettingsService", "getNumericValue should coerce values correctly", function() {
    setupMockSheet();
    TestSettingsService.setValue("num_10", 10);
    TestSettingsService.setValue("num_float", 12.34);
    TestSettingsService.setValue("string_num", "56.7");
    TestSettingsService.setValue("string_invalid", "abc");

    T.assertEquals(10, TestSettingsService.getNumericValue("num_10", 0), "Number 10 should return 10.");
    T.assertEquals(12.34, TestSettingsService.getNumericValue("num_float", 0), "Number 12.34 should return 12.34.");
    T.assertEquals(56.7, TestSettingsService.getNumericValue("string_num", 0), "String '56.7' should return 56.7.");
    T.assertEquals(99, TestSettingsService.getNumericValue("string_invalid", 99), "Invalid string should return default (99).");
    T.assertEquals(55, TestSettingsService.getNumericValue("nonexistent", 55), "Nonexistent key should return default (55).");
  });

  T.registerTest("SettingsService", "toggleBooleanValue should toggle correctly", function() {
    setupMockSheet();
    const key = "toggle_test";

    // Test 1: Default is false, toggle to true
    let newValue = TestSettingsService.toggleBooleanValue(key, false);
    T.assertTrue(newValue, "Toggle 1: Should return true.");
    T.assertTrue(TestSettingsService.getBooleanValue(key, false), "Toggle 1: Stored value should be true.");

    // Test 2: Toggle existing true to false
    newValue = TestSettingsService.toggleBooleanValue(key, false);
    T.assertFalse(newValue, "Toggle 2: Should return false.");
    T.assertFalse(TestSettingsService.getBooleanValue(key, true), "Toggle 2: Stored value should be false.");

     // Test 3: Toggle existing false to true
    newValue = TestSettingsService.toggleBooleanValue(key, false);
    T.assertTrue(newValue, "Toggle 3: Should return true.");
    T.assertTrue(TestSettingsService.getBooleanValue(key, false), "Toggle 3: Stored value should be true.");
  });
  
  T.registerTest("SettingsService", "ShowSubCategories methods should work together", function() {
      setupMockSheet();
      // Default should be true
      T.assertTrue(TestSettingsService.getShowSubCategories(), "Default ShowSubCategories should be true.");
      
      // Toggle to false
      TestSettingsService.toggleShowSubCategories();
      T.assertFalse(TestSettingsService.getShowSubCategories(), "ShowSubCategories should be false after toggle.");
      
      // Set explicitly to true
      TestSettingsService.setShowSubCategories(true);
      T.assertTrue(TestSettingsService.getShowSubCategories(), "ShowSubCategories should be true after explicit set.");
  });

  T.registerTest("SettingsService", "getAllPreferences should return all settings", function() {
    setupMockSheet();
    TestSettingsService.setValue("pref1", "value1");
    TestSettingsService.setValue("pref2", true);
    TestSettingsService.setValue("pref3", 123);

    const allPrefs = TestSettingsService.getAllPreferences();
    T.assertEquals("value1", allPrefs["pref1"], "getAllPreferences: pref1 should exist.");
    T.assertEquals(true, allPrefs["pref2"], "getAllPreferences: pref2 should exist.");
    T.assertEquals(123, allPrefs["pref3"], "getAllPreferences: pref3 should exist.");
    T.assertEquals(3, Object.keys(allPrefs).length, "getAllPreferences: Should contain 3 preferences.");
  });

  T.registerTest("SettingsService", "resetAllPreferences should clear settings", function() {
    setupMockSheet();
    TestSettingsService.setValue("pref1", "value1");
    TestSettingsService.setValue("pref2", true);

    TestSettingsService.resetAllPreferences();

    const allPrefs = TestSettingsService.getAllPreferences();
    T.assertEquals(0, Object.keys(allPrefs).length, "getAllPreferences: Should be empty after reset.");
    T.assertNotNull(lastSuccessNotification, "Success notification should be shown after reset.");
    // Check sheet data directly
    const sheetData = mockSheetData[mockConfig._settingsSheetName];
    T.assertEquals(1, sheetData.length, "Sheet should only contain header row after reset.");
  });


})(); // End IIFE
