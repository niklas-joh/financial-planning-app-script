/**
 * Financial Planning Tools - Settings Service Tests
 *
 * Exercises the real FinancialPlanner.SettingsService against the in-memory
 * Apps Script fakes in gas-fakes.js. Preferences live on a hidden "Settings"
 * sheet, so assertions read that sheet back directly.
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Settings = FinancialPlanner.SettingsService;
  let gas;

  function setupMockSheet() {
    gas = FakeGas.install(global);
  }

  function settingsSheet() {
    return gas.spreadsheet.getSheetByName(FinancialPlanner.Config.getSheetNames().SETTINGS);
  }

  T.registerTest("SettingsService", "creates a hidden Settings sheet with a header on first use", function() {
    setupMockSheet();
    Settings.getValue("anything", null);
    const sheet = settingsSheet();
    T.assertNotNull(sheet, "Settings sheet should be created.");
    T.assertTrue(sheet.isSheetHidden(), "Settings sheet should be hidden.");
    T.assertDeepEquals([["Preference", "Value"]], sheet._values(), "Header row should be written.");
  });

  T.registerTest("SettingsService", "setPlaidEnvironment rejects unknown environments", function() {
    setupMockSheet();
    T.assertThrows(function() { Settings.setPlaidEnvironment("staging"); }, "Invalid environment");
    T.assertEquals("sandbox", Settings.getPlaidEnvironment(), "Default environment should stay sandbox.");
    Settings.setPlaidEnvironment("production");
    T.assertEquals("production", Settings.getPlaidEnvironment(), "Valid environment should be stored.");
  });

  T.registerTest("SettingsService", "getValue should return default value if key not found", function() {
    setupMockSheet();
    const defaultValue = "default";
    const value = Settings.getValue("nonexistent_key", defaultValue);
    T.assertEquals(defaultValue, value, "Should return default value for nonexistent key.");
  });

  T.registerTest("SettingsService", "setValue should add a new preference", function() {
    setupMockSheet();
    const key = "new_pref";
    const value = "new_value";
    Settings.setValue(key, value);
    const retrievedValue = Settings.getValue(key, "default");
    T.assertEquals(value, retrievedValue, "Should retrieve the newly set value.");
    const sheetData = settingsSheet()._values();
    T.assertEquals(2, sheetData.length, "Sheet should have header + 1 data row.");
    T.assertEquals(key, sheetData[1][0], "Key should be in the second row, first column.");
    T.assertEquals(value, sheetData[1][1], "Value should be in the second row, second column.");
  });

  T.registerTest("SettingsService", "setValue should update an existing preference", function() {
    setupMockSheet();
    const key = "existing_pref";
    const initialValue = "initial_value";
    const updatedValue = "updated_value";
    Settings.getValue(key, null); // creates the Settings sheet
    settingsSheet().appendRow([key, initialValue]);

    Settings.setValue(key, updatedValue);
    const retrievedValue = Settings.getValue(key, "default");
    T.assertEquals(updatedValue, retrievedValue, "Should retrieve the updated value.");
    const sheetData = settingsSheet()._values();
    T.assertEquals(2, sheetData.length, "Sheet should still have header + 1 data row.");
    T.assertEquals(key, sheetData[1][0], "Key should be in the second row, first column.");
    T.assertEquals(updatedValue, sheetData[1][1], "Value should be updated in the second row, second column.");
  });

  T.registerTest("SettingsService", "getBooleanValue should coerce values correctly", function() {
    setupMockSheet();
    Settings.setValue("bool_true", true);
    Settings.setValue("bool_false", false);
    Settings.setValue("string_true", "true");
    Settings.setValue("string_false", "false");
    Settings.setValue("num_1", 1);
    Settings.setValue("num_0", 0);
    Settings.setValue("string_1", "1");
    Settings.setValue("string_0", "0");
    Settings.setValue("other_string", "hello");

    T.assertTrue(Settings.getBooleanValue("bool_true", false), "Boolean true should return true.");
    T.assertFalse(Settings.getBooleanValue("bool_false", true), "Boolean false should return false.");
    T.assertTrue(Settings.getBooleanValue("string_true", false), "String 'true' should return true.");
    T.assertFalse(Settings.getBooleanValue("string_false", true), "String 'false' should return false.");
    T.assertTrue(Settings.getBooleanValue("num_1", false), "Number 1 should return true.");
    T.assertFalse(Settings.getBooleanValue("num_0", true), "Number 0 should return false.");
     T.assertTrue(Settings.getBooleanValue("string_1", false), "String '1' should return true.");
    T.assertFalse(Settings.getBooleanValue("string_0", true), "String '0' should return false.");
    T.assertFalse(Settings.getBooleanValue("other_string", false), "Other string should return default (false).");
    T.assertTrue(Settings.getBooleanValue("nonexistent", true), "Nonexistent key should return default (true).");
  });

  T.registerTest("SettingsService", "getNumericValue should coerce values correctly", function() {
    setupMockSheet();
    Settings.setValue("num_10", 10);
    Settings.setValue("num_float", 12.34);
    Settings.setValue("string_num", "56.7");
    Settings.setValue("string_invalid", "abc");

    T.assertEquals(10, Settings.getNumericValue("num_10", 0), "Number 10 should return 10.");
    T.assertEquals(12.34, Settings.getNumericValue("num_float", 0), "Number 12.34 should return 12.34.");
    T.assertEquals(56.7, Settings.getNumericValue("string_num", 0), "String '56.7' should return 56.7.");
    T.assertEquals(99, Settings.getNumericValue("string_invalid", 99), "Invalid string should return default (99).");
    T.assertEquals(55, Settings.getNumericValue("nonexistent", 55), "Nonexistent key should return default (55).");
  });

  T.registerTest("SettingsService", "toggleBooleanValue should toggle correctly", function() {
    setupMockSheet();
    const key = "toggle_test";

    // Test 1: Default is false, toggle to true
    let newValue = Settings.toggleBooleanValue(key, false);
    T.assertTrue(newValue, "Toggle 1: Should return true.");
    T.assertTrue(Settings.getBooleanValue(key, false), "Toggle 1: Stored value should be true.");

    // Test 2: Toggle existing true to false
    newValue = Settings.toggleBooleanValue(key, false);
    T.assertFalse(newValue, "Toggle 2: Should return false.");
    T.assertFalse(Settings.getBooleanValue(key, true), "Toggle 2: Stored value should be false.");

     // Test 3: Toggle existing false to true
    newValue = Settings.toggleBooleanValue(key, false);
    T.assertTrue(newValue, "Toggle 3: Should return true.");
    T.assertTrue(Settings.getBooleanValue(key, false), "Toggle 3: Stored value should be true.");
  });
  
  T.registerTest("SettingsService", "ShowSubCategories methods should work together", function() {
      setupMockSheet();
      // Default should be true
      T.assertTrue(Settings.getShowSubCategories(), "Default ShowSubCategories should be true.");
      
      // Toggle to false
      Settings.toggleShowSubCategories();
      T.assertFalse(Settings.getShowSubCategories(), "ShowSubCategories should be false after toggle.");
      
      // Set explicitly to true
      Settings.setShowSubCategories(true);
      T.assertTrue(Settings.getShowSubCategories(), "ShowSubCategories should be true after explicit set.");
  });

  T.registerTest("SettingsService", "getAllPreferences should return all settings", function() {
    setupMockSheet();
    Settings.setValue("pref1", "value1");
    Settings.setValue("pref2", true);
    Settings.setValue("pref3", 123);

    const allPrefs = Settings.getAllPreferences();
    T.assertEquals("value1", allPrefs["pref1"], "getAllPreferences: pref1 should exist.");
    T.assertEquals(true, allPrefs["pref2"], "getAllPreferences: pref2 should exist.");
    T.assertEquals(123, allPrefs["pref3"], "getAllPreferences: pref3 should exist.");
    T.assertEquals(3, Object.keys(allPrefs).length, "getAllPreferences: Should contain 3 preferences.");
  });

  T.registerTest("SettingsService", "resetAllPreferences should clear settings", function() {
    setupMockSheet();
    Settings.setValue("pref1", "value1");
    Settings.setValue("pref2", true);

    Settings.resetAllPreferences();

    const allPrefs = Settings.getAllPreferences();
    T.assertEquals(0, Object.keys(allPrefs).length, "getAllPreferences: Should be empty after reset.");
    T.assertEquals('All preferences have been reset.', gas.spreadsheet.toasts.pop().message, "Success notification should be shown after reset.");
    const sheetData = settingsSheet()._values();
    T.assertEquals(1, sheetData.length, "Sheet should only contain header row after reset.");
  });


})();
