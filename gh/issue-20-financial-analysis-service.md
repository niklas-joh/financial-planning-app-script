# Refactor Analytics into Dedicated Service

## Current Implementation Analysis

Looking at the code in `finance_overview.js`, there are several functions that handle financial metrics, expense categories analysis, and chart creation:

1. `addKeyMetricsSection()`
2. `addExpenseCategoriesSection()`
3. `findExpenseCategories()`
4. `createExpenditureCharts()`

These functions are currently:
- Called from the `addMetrics()` method in the `FinancialOverviewBuilder` class
- Using a `startRow` parameter to position content relative to other data
- Directly modifying the Overview sheet instead of a dedicated Analysis sheet
- Structured procedurally rather than using object-oriented design

## Issues with the Current Implementation

The current approach has several drawbacks:

1. **Maintainability**: The functions are tightly coupled and don't follow a clear object-oriented structure
2. **Positioning**: Using `startRow` makes the layout dependent on other content
3. **Separation of Concerns**: Analysis content is mixed with the overview sheet
4. **Extensibility**: Adding new metrics or visualizations requires modifying existing functions

## Proposed Solution

I recommend creating a dedicated `FinancialAnalysisService` class in a new file that would:

1. Create and manage a dedicated "Analysis" sheet
2. Encapsulate all analytics-related functionality
3. Use object-oriented design for better maintainability
4. Decouple from the positioning constraints of the overview sheet

## Implementation Plan

1. Create a new file `src/features/financial_analysis.js`
2. Design a `FinancialAnalysisService` class with methods for:
   - Creating/accessing the Analysis sheet
   - Generating key metrics
   - Building category breakdowns
   - Creating visualizations
3. Modify the `FinancialOverviewBuilder` to use this service
4. Add configuration for the Analysis sheet in the main config object

## Proposed Class Structure

```javascript
class FinancialAnalysisService {
  constructor(spreadsheet, overviewSheet, config) {
    this.spreadsheet = spreadsheet;
    this.overviewSheet = overviewSheet;
    this.config = config;
    this.analysisSheet = this.getOrCreateAnalysisSheet();
    this.data = null;
    this.totals = null;
  }

  // Core methods
  initialize() {}
  analyze() {}
  
  // Sheet management
  getOrCreateAnalysisSheet() {}
  clearSheet() {}
  
  // Key metrics section
  addKeyMetricsSection() {}
  
  // Expense categories
  addExpenseCategoriesSection() {}
  findExpenseCategories() {}
  
  // Visualization
  createExpenditureCharts() {}
}
```

## Tasks
- [ ] Create new `financial_analysis.js` file
- [ ] Implement `FinancialAnalysisService` class
- [ ] Move analytics functions (`addKeyMetricsSection`, `addExpenseCategoriesSection`, etc.)
- [ ] Update `FinancialOverviewBuilder` to use the new service
- [ ] Add configuration for the Analysis sheet
- [ ] Test to ensure all functionality works as expected

## Benefits
- Improved code organization and maintainability
- Better separation of concerns
- More flexible positioning of analytics content
- Easier to extend with new analytics features
