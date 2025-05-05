/**
 * Test script for sub-category toggle functionality
 */
function testSubCategoryToggle() {
  // Setup test environment
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const overviewSheet = ss.getSheetByName("Overview");
  
  // Test 1: Verify initial state (sub-categories should be shown by default)
  const initialShowSubCategories = getUserPreference("ShowSubCategories", true);
  Logger.log("Initial Sub-Categories Preference: " + initialShowSubCategories);
  
  // Regenerate overview to ensure current state
  createFinancialOverview();
  
  // Test 2: Check checkbox state matches preference
  const checkboxCell = overviewSheet.getRange("S1");
  const checkboxValue = checkboxCell.getValue();
  
  if (checkboxValue !== initialShowSubCategories) {
    throw new Error("Checkbox state does not match user preference");
  }
  
  // Test 3: Toggle sub-categories and verify
  toggleShowSubCategories();
  
  const newShowSubCategories = getUserPreference("ShowSubCategories", true);
  Logger.log("New Sub-Categories Preference: " + newShowSubCategories);
  
  if (newShowSubCategories === initialShowSubCategories) {
    throw new Error("Sub-category toggle did not change preference");
  }
  
  // Test 4: Verify overview regeneration
  createFinancialOverview();
  
  // Additional checks can be added here to verify specific behavior
  
  Logger.log("All sub-category toggle tests passed successfully!");
}
