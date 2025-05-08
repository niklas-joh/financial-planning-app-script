/**
 * Financial Planning Tools - Financial Analysis Service
 * 
 * This module provides analytics functionality for financial data through a dedicated
 * service. It creates a separate Analysis sheet with key metrics, expense category analysis,
 * and visualizations.
 * 
 * Version: 2.1.0
 * Last Updated: 2025-05-08
 */

/**
 * @namespace FinancialPlanner.FinancialAnalysisService
 * @description Service for performing financial analysis based on the data aggregated in the 'Overview' sheet.
 * It generates key metrics, analyzes expense categories against targets, and creates visualizations in a dedicated 'Analysis' sheet.
 * @param {FinancialPlanner.Utils} utils - The utility service.
 * @param {FinancialPlanner.UIService} uiService - The UI service for notifications.
 * @param {FinancialPlanner.ErrorService} errorService - The error handling service.
 * @param {FinancialPlanner.Config} config - The global configuration service.
 */
FinancialPlanner.FinancialAnalysisService = (function(utils, uiService, errorService, config) {
  // ============================================================================
  // PRIVATE IMPLEMENTATION
  // ============================================================================
  
  /**
   * Internal class responsible for performing the financial analysis calculations and sheet manipulations.
   * @class FinancialAnalysisService
   * @private
   */
  class FinancialAnalysisService {
    /**
     * Creates an instance of the internal FinancialAnalysisService class.
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet object.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} overviewSheet - The sheet object containing the generated financial overview data.
     * @param {object} analysisConfig - A configuration object, typically derived from `FinancialPlanner.Config.get()`.
     * @constructor
     */
    constructor(spreadsheet, overviewSheet, analysisConfig) {
      /**
       * The active spreadsheet.
       * @type {GoogleAppsScript.Spreadsheet.Spreadsheet}
       * @private
       */
      this.spreadsheet = spreadsheet;
      
      /**
       * The overview sheet containing summarized financial data.
       * @type {GoogleAppsScript.Spreadsheet.Sheet}
       * @private
       */
      this.overviewSheet = overviewSheet;
      
      /**
       * Configuration object used by the analysis service.
       * @type {object}
       * @private
       */
      this.config = analysisConfig;
      
      /**
       * The sheet where analysis results are displayed. Created if it doesn't exist.
       * @type {GoogleAppsScript.Spreadsheet.Sheet}
       * @private
       */
      this.analysisSheet = utils.getOrCreateSheet(spreadsheet, this.config.SHEETS.ANALYSIS);
      
      /**
       * Holds the data extracted from the overview sheet for analysis.
       * Populated by `extractDataFromOverview`.
       * @type {{incomeCategories: Array<object>, expenseCategories: Array<object>, savingsCategories: Array<object>, months: Array<string>} | null}
       * @private
       */
      this.data = null;
      
      /**
       * Holds calculated total values (average monthly) for key categories (Income, Expenses, Savings, etc.).
       * Populated by `extractDataFromOverview`.
       * @type {{income: {row: number, value: number}, expenses: {row: number, value: number}, savings: {row: number, value: number}, essentials: {row: number, value: number}, wantsPleasure: {row: number, value: number}, extra: {row: number, value: number}} | null}
       * @private
       */
      this.totals = null;
    }

    /**
     * Initializes the analysis service instance.
     * Extracts data from the overview sheet and sets up the analysis sheet structure and formatting.
     * @return {void}
     * @public
     */
    initialize() {
      // Extract data from the overview sheet for analysis
      this.extractDataFromOverview();
      
      // Clear and set up the analysis sheet
      this.setupAnalysisSheet();
    }

    /**
     * Executes the core analysis workflow.
     * Calls methods to add key metrics, expense category analysis, and charts to the analysis sheet.
     * @return {void}
     * @public
     */
    analyze() {
      // Start at row 2 (after header)
      let currentRow = 2;
      
      // Add key metrics section
      currentRow = this.addKeyMetricsSection(currentRow);
      
      // Add space between sections
      currentRow += 2;
      
      // Add expense categories section
      currentRow = this.addExpenseCategoriesSection(currentRow);
      
      // Add space between sections
      currentRow += 2;
      
      // Create expenditure charts
      this.createExpenditureCharts(currentRow);
    }

    /**
     * Extracts and structures relevant data from the 'Overview' sheet.
     * Populates `this.data` with categorized amounts and `this.totals` with key summary figures (like total income, expenses).
     * It specifically looks for rows starting with "Total [Type]" to get summary values.
     * @return {void}
     * @private
     */
    extractDataFromOverview() {
      // Get all data from the overview sheet
      const overviewData = this.overviewSheet.getDataRange().getValues();
      
      // Initialize data structure
      this.data = {
        incomeCategories: [],
        expenseCategories: [],
        savingsCategories: [],
        months: []
      };
      
      // Initialize totals
      this.totals = {
        income: { row: -1, value: 0 },
        expenses: { row: -1, value: 0 },
        savings: { row: -1, value: 0 },
        // Add expense type totals
        essentials: { row: -1, value: 0 },
        wantsPleasure: { row: -1, value: 0 },
        extra: { row: -1, value: 0 }
      };
      
      // Extract month names from headers (columns 5-16 in overview sheet)
      for (let i = 4; i <= 15; i++) {
        this.data.months.push(overviewData[0][i]);
      }
      
      // Find rows containing total Income, Expenses, and Savings
      for (let i = 0; i < overviewData.length; i++) {
        const rowData = overviewData[i];
        
        // Check for total rows
        if (rowData[0] === "Total Income") {
          this.totals.income.row = i + 1;
          this.totals.income.value = rowData[16]; // Average column
        } else if (rowData[0] === "Total Essentials") {
          // Track Essentials separately
          this.totals.essentials.row = i + 1;
          this.totals.essentials.value = rowData[16]; // Average column
          
          // Also add to total expenses
          if (this.totals.expenses.row === -1) {
            this.totals.expenses.row = i + 1;
            this.totals.expenses.value = 0;
          }
          this.totals.expenses.value += rowData[16];
        } else if (rowData[0] === "Total Wants/Pleasure") {
          // Track Wants/Pleasure separately
          this.totals.wantsPleasure.row = i + 1;
          this.totals.wantsPleasure.value = rowData[16]; // Average column
          
          // Also add to total expenses
          if (this.totals.expenses.row === -1) {
            this.totals.expenses.row = i + 1;
            this.totals.expenses.value = 0;
          }
          this.totals.expenses.value += rowData[16];
        } else if (rowData[0] === "Total Extra") {
          // Track Extra separately
          this.totals.extra.row = i + 1;
          this.totals.extra.value = rowData[16]; // Average column
          
          // Also add to total expenses
          if (this.totals.expenses.row === -1) {
            this.totals.expenses.row = i + 1;
            this.totals.expenses.value = 0;
          }
          this.totals.expenses.value += rowData[16];
        } else if (rowData[0] === "Total Savings") {
          this.totals.savings.row = i + 1;
          this.totals.savings.value = rowData[16]; // Average column
        }
        
        // Extract categories
        if (rowData[0] === "Income" && rowData[1]) {
          this.data.incomeCategories.push({
            category: rowData[1],
            subcategory: rowData[2] || "",
            amount: rowData[16], // Average column
            row: i + 1
          });
        } else if ((rowData[0] === "Essentials" || rowData[0] === "Wants/Pleasure" || rowData[0] === "Extra") && rowData[1]) {
          this.data.expenseCategories.push({
            type: rowData[0],
            category: rowData[1],
            subcategory: rowData[2] || "",
            amount: rowData[16], // Average column
            row: i + 1
          });
        } else if (rowData[0] === "Savings" && rowData[1]) {
          this.data.savingsCategories.push({
            category: rowData[1],
            subcategory: rowData[2] || "",
            amount: rowData[16], // Average column
            row: i + 1
          });
        }
      }
    }

    /**
     * Clears and sets up the basic structure and formatting of the 'Analysis' sheet.
     * Includes setting the main header, freezing the header row, and setting initial column widths.
     * @return {void}
     * @private
     */
    setupAnalysisSheet() {
      // Clear existing content
      this.analysisSheet.clear();
      this.analysisSheet.clearFormats();
      
      // Set up header and formatting
      this.analysisSheet.getRange("A1").setValue("Financial Analysis");
      this.analysisSheet.getRange("A1:J1")
        .setBackground(this.config.COLORS.UI.HEADER_BG)
        .setFontWeight("bold")
        .setFontColor(this.config.COLORS.UI.HEADER_FONT);
      
      // Freeze the header row
      this.analysisSheet.setFrozenRows(1);
      
      // Set column widths for better readability
      this.analysisSheet.setColumnWidth(1, 200); // Metric/Category
      this.analysisSheet.setColumnWidth(2, 120); // Value
      this.analysisSheet.setColumnWidth(3, 120); // Target
      
      // Set sheet description
      this.analysisSheet.setName(this.config.SHEETS.ANALYSIS);
    }

    /**
     * Adds the 'Key Metrics' section to the analysis sheet.
     * Calculates and displays metrics like Expenses/Income Ratio, Savings Rate, and individual expense category rates against targets.
     * Applies formatting and conditional formatting for readability.
     * @param {number} startRow - The 1-based row index where the section should start.
     * @return {number} The next available row index after adding the section.
     * @private
     */
    addKeyMetricsSection(startRow) {
      // Add Key Metrics header
      this.analysisSheet.getRange(startRow, 1).setValue("Key Metrics");
      this.analysisSheet.getRange(startRow, 1, 1, 4) // Expanded to include description column
        .setBackground(this.config.COLORS.UI.HEADER_BG)
        .setFontWeight("bold")
        .setFontColor(this.config.COLORS.UI.HEADER_FONT)
        .setHorizontalAlignment("center");
      
      startRow++;
      
      // Create a metrics table
      this.analysisSheet.getRange(startRow, 1, 1, 4) // Expanded to include description column
        .setValues([["Metric", "Value", "Target", "Description"]])
        .setBackground("#F5F5F5")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      
      // Set width for description column
      this.analysisSheet.setColumnWidth(4, 300);
      
      startRow++;
      
      // Initialize arrays for batch processing
      const metricsStartRow = startRow;
      let currentMetricRow = 0;
      
      // Arrays for batch processing
      const metricValues = [];
      const metricFormulas = [];
      const metricTargets = [];
      const metricDescriptions = [];
      const metricBackgrounds = [];
      const conditionalFormatRules = [];
      
      // Consistent background color for all metrics
      const metricBgColor = this.config.COLORS.UI.METRICS_BG;
      
      // 1. Expenses/Income Ratio
      if (this.totals.income.row > 0 && this.totals.expenses.row > 0) {
        metricValues.push(["Expenses/Income Ratio"]);
        metricFormulas.push([`=-${this.totals.expenses.value}/${this.totals.income.value}`]);
        metricTargets.push([this.config.TARGET_RATES.DEFAULT * -1]); // Use config target rate with negative sign
        metricDescriptions.push(["Negative % indicates spending, lower absolute value is better"]);
        metricBackgrounds.push([metricBgColor]);
        
        // Add conditional formatting (green if meeting target, red if not)
        // Use direct cell reference instead of string formula
        const targetCell = this.analysisSheet.getRange(startRow + currentMetricRow, 3);
        const valueCell = this.analysisSheet.getRange(startRow + currentMetricRow, 2);
        
        conditionalFormatRules.push({
          row: startRow + currentMetricRow,
          rule: SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`B${startRow + currentMetricRow}<C${startRow + currentMetricRow}`)
            .setBackground("#FFCDD2") // Light red if below target (more negative)
            .setRanges([valueCell])
        });
        
        currentMetricRow++;
      }
      
      // 2. Essentials Rate
      if (this.totals.income.row > 0 && this.totals.essentials.row > 0) {
        metricValues.push(["Essentials Rate"]);
        metricFormulas.push([`=${this.totals.essentials.value}/${this.totals.income.value}`]);
        metricTargets.push([this.config.TARGET_RATES.ESSENTIALS]);
        metricDescriptions.push(["Percentage of income spent on essential expenses (lower is better)"]);
        metricBackgrounds.push([metricBgColor]);
        
        // Add conditional formatting (red if exceeding target)
        const targetCell = this.analysisSheet.getRange(startRow + currentMetricRow, 3);
        const valueCell = this.analysisSheet.getRange(startRow + currentMetricRow, 2);
        
        conditionalFormatRules.push({
          row: startRow + currentMetricRow,
          rule: SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`B${startRow + currentMetricRow}>C${startRow + currentMetricRow}`)
            .setBackground("#FFCDD2") // Light red if above target
            .setRanges([valueCell])
        });
        
        currentMetricRow++;
      }
      
      // 3. Wants/Pleasure Rate
      if (this.totals.income.row > 0 && this.totals.wantsPleasure.row > 0) {
        metricValues.push(["Wants/Pleasure Rate"]);
        metricFormulas.push([`=${this.totals.wantsPleasure.value}/${this.totals.income.value}`]);
        metricTargets.push([this.config.TARGET_RATES.WANTS_PLEASURE]);
        metricDescriptions.push(["Percentage of income spent on wants and pleasure (discretionary spending)"]);
        metricBackgrounds.push([metricBgColor]);
        
        // Add conditional formatting (red if exceeding target)
        const targetCell = this.analysisSheet.getRange(startRow + currentMetricRow, 3);
        const valueCell = this.analysisSheet.getRange(startRow + currentMetricRow, 2);
        
        conditionalFormatRules.push({
          row: startRow + currentMetricRow,
          rule: SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`B${startRow + currentMetricRow}>C${startRow + currentMetricRow}`)
            .setBackground("#FFCDD2") // Light red if above target
            .setRanges([valueCell])
        });
        
        currentMetricRow++;
      }
      
      // 4. Extra Expenses Rate
      if (this.totals.income.row > 0 && this.totals.extra.row > 0) {
        metricValues.push(["Extra Expenses Rate"]);
        metricFormulas.push([`=${this.totals.extra.value}/${this.totals.income.value}`]);
        metricTargets.push([this.config.TARGET_RATES.EXTRA]);
        metricDescriptions.push(["Percentage of income spent on extra/miscellaneous expenses"]);
        metricBackgrounds.push([metricBgColor]);
        
        // Add conditional formatting (red if exceeding target)
        const targetCell = this.analysisSheet.getRange(startRow + currentMetricRow, 3);
        const valueCell = this.analysisSheet.getRange(startRow + currentMetricRow, 2);
        
        conditionalFormatRules.push({
          row: startRow + currentMetricRow,
          rule: SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`B${startRow + currentMetricRow}>C${startRow + currentMetricRow}`)
            .setBackground("#FFCDD2") // Light red if above target
            .setRanges([valueCell])
        });
        
        currentMetricRow++;
      }
      
      // Add separator (empty row with light gray background)
      if (currentMetricRow > 0) {
        metricValues.push([""]);
        metricFormulas.push([""]);
        metricTargets.push([""]);
        metricDescriptions.push([""]);
        metricBackgrounds.push(["#E0E0E0"]); // Light gray separator
        
        // Add a horizontal line for visual separation
        this.analysisSheet.getRange(startRow + currentMetricRow, 1, 1, 4)
          .setBorder(false, false, true, false, false, false, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID);
        
        currentMetricRow++;
      }
      
      // 5. Savings Rate
      if (this.totals.income.row > 0 && this.totals.savings.row > 0) {
        metricValues.push(["Savings Rate"]);
        metricFormulas.push([`=-${this.totals.savings.value}/${this.totals.income.value}`]);
        metricTargets.push([this.config.TARGET_RATES.DEFAULT]); // Use config target rate
        metricDescriptions.push(["Positive % indicates saving money, negative % indicates withdrawing from savings"]);
        metricBackgrounds.push([metricBgColor]);
        
        // Add conditional formatting (green if meeting target, red if not)
        const targetCell = this.analysisSheet.getRange(startRow + currentMetricRow, 3);
        const valueCell = this.analysisSheet.getRange(startRow + currentMetricRow, 2);
        
        conditionalFormatRules.push({
          row: startRow + currentMetricRow,
          rule: SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`B${startRow + currentMetricRow}<C${startRow + currentMetricRow}`)
            .setBackground("#FFCDD2") // Light red if below target
            .setRanges([valueCell])
        });
        
        currentMetricRow++;
      }
      
      // Add another separator
      if (currentMetricRow > 0) {
        metricValues.push([""]);
        metricFormulas.push([""]);
        metricTargets.push([""]);
        metricDescriptions.push([""]);
        metricBackgrounds.push(["#E0E0E0"]); // Light gray separator
        
        // Add a horizontal line for visual separation
        this.analysisSheet.getRange(startRow + currentMetricRow, 1, 1, 4)
          .setBorder(false, false, true, false, false, false, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID);
        
        currentMetricRow++;
      }
      
      // 6. Monthly Net Cash Flow
      if (this.totals.income.row > 0 && this.totals.expenses.row > 0) {
        metricValues.push(["Monthly Net Cash Flow"]);
        metricFormulas.push([`=${this.totals.income.value}-${this.totals.expenses.value}`]);
        metricTargets.push([0]); // Target is positive cash flow
        metricDescriptions.push(["Positive value means you're earning more than spending, negative means you're spending more than earning"]);
        metricBackgrounds.push([metricBgColor]);
        
        // Add conditional formatting (green if positive, red if negative)
        const valueCell = this.analysisSheet.getRange(startRow + currentMetricRow, 2);
        
        conditionalFormatRules.push({
          row: startRow + currentMetricRow,
          rule: SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`B${startRow + currentMetricRow}<0`)
            .setBackground("#FFCDD2") // Light red if negative
            .setRanges([valueCell])
        });
        
        currentMetricRow++;
      }
      
      // Write all data to the sheet in batches if we have metrics
      if (currentMetricRow > 0) {
        // Set metric names
        if (metricValues.length > 0) {
          this.analysisSheet.getRange(startRow, 1, metricValues.length, 1).setValues(metricValues);
        }
        
        // Set formulas
        if (metricFormulas.length > 0) {
          this.analysisSheet.getRange(startRow, 2, metricFormulas.length, 1).setFormulas(metricFormulas);
        }
        
        // Set targets
        if (metricTargets.length > 0) {
          this.analysisSheet.getRange(startRow, 3, metricTargets.length, 1).setValues(metricTargets);
        }
        
        // Set descriptions
        if (metricDescriptions.length > 0) {
          this.analysisSheet.getRange(startRow, 4, metricDescriptions.length, 1).setValues(metricDescriptions);
        }
        
        // Set backgrounds
        if (metricBackgrounds.length > 0) {
          this.analysisSheet.getRange(startRow, 1, metricBackgrounds.length, 4).setBackgrounds(
            metricBackgrounds.map(bg => [bg[0], bg[0], bg[0], bg[0]])
          );
        }
        
        // Format percentage cells
        const percentageRows = metricValues.map((_, i) => startRow + i).filter(row => 
          this.analysisSheet.getRange(row, 1).getValue() !== "Monthly Net Cash Flow" && 
          this.analysisSheet.getRange(row, 1).getValue() !== ""
        );
        
        if (percentageRows.length > 0) {
          percentageRows.forEach(row => {
            utils.formatAsPercentage(this.analysisSheet.getRange(row, 2, 1, 2));
          });
        }
        
        // Format currency cells
        const currencyRows = metricValues.map((_, i) => startRow + i).filter(row => 
          this.analysisSheet.getRange(row, 1).getValue() === "Monthly Net Cash Flow"
        );
        
        if (currencyRows.length > 0) {
          currencyRows.forEach(row => {
            utils.formatAsCurrency(this.analysisSheet.getRange(row, 2, 1, 2));
          });
        }
        
        // Apply conditional formatting rules
        if (conditionalFormatRules.length > 0) {
          const rules = this.analysisSheet.getConditionalFormatRules();
          conditionalFormatRules.forEach(item => {
            rules.push(item.rule);
          });
          this.analysisSheet.setConditionalFormatRules(rules);
        }
        
        // Add a border around the metrics table
        this.analysisSheet.getRange(metricsStartRow, 1, currentMetricRow, 4).setBorder(
          true, true, true, true, true, true, 
          "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID
        );
        
        // Add subtle shading to every other non-separator row for better readability
        const nonSeparatorRows = metricValues.map((val, i) => ({
          row: startRow + i,
          isEmpty: val[0] === ""
        })).filter(item => !item.isEmpty);
        
        nonSeparatorRows.forEach((item, index) => {
          if (index % 2 === 1) {
            this.analysisSheet.getRange(item.row, 1, 1, 4)
              .setBackground("#F8F8F8"); // Very light gray for alternate rows
          }
        });
      }
      
      return startRow + currentMetricRow;
    }

    /**
     * Adds the 'Expense Categories' section to the analysis sheet.
     * Lists major expense categories (excluding sub-categories), their average monthly amount,
     * percentage of income, target percentage, and variance.
     * Applies formatting and conditional formatting.
     * @param {number} startRow - The 1-based row index where the section should start.
     * @return {number} The next available row index after adding the section.
     * @private
     */
    addExpenseCategoriesSection(startRow) {
      // Add Expense Categories header
      this.analysisSheet.getRange(startRow, 1).setValue("Expense Categories");
      this.analysisSheet.getRange(startRow, 1, 1, 6)
        .setBackground(this.config.COLORS.UI.HEADER_BG)
        .setFontWeight("bold")
        .setFontColor(this.config.COLORS.UI.HEADER_FONT)
        .setHorizontalAlignment("center");
      
      startRow++;
      
      // Add headers
      this.analysisSheet.getRange(startRow, 1, 1, 6)
        .setValues([["Category", "Type", "Amount", "% of Income", "Target %", "Variance"]])
        .setBackground("#F5F5F5")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      
      startRow++;
      
      // Initialize arrays for batch processing
      const categoryStartRow = startRow;
      let currentCategoryRow = 0;
      
      // Arrays for batch processing
      const categoryValues = [];
      const typeValues = [];
      const amountValues = [];
      const percentFormulas = [];
      const targetValues = [];
      const varianceFormulas = [];
      const backgroundColors = [];
      const conditionalFormatRules = [];
      
      // Consistent background color
      const categoryBgColor = this.config.COLORS.UI.METRICS_BG;
      
      // Add rows for each expense category
      if (this.data.expenseCategories.length > 0) {
        // Sort categories by amount (descending)
        const sortedCategories = [...this.data.expenseCategories]
          .filter(category => !category.subcategory) // Skip subcategories
          .sort((a, b) => b.amount - a.amount);
        
        // Prepare data for batch processing
        sortedCategories.forEach((category) => {
          categoryValues.push([category.category]);
          typeValues.push([category.type]);
          amountValues.push([category.amount]);
          
          // Calculate percentage of income
          if (this.totals.income.value > 0) {
            percentFormulas.push([`=C${startRow + currentCategoryRow}/${this.totals.income.value}`]);
          } else {
            percentFormulas.push([0]);
          }
          
          // Set target rate based on expense type
          let targetRate = this.config.TARGET_RATES.DEFAULT; // Default
          if (category.type === "Essentials") {
            targetRate = this.config.TARGET_RATES.ESSENTIALS;
          } else if (category.type === "Wants/Pleasure") {
            targetRate = this.config.TARGET_RATES.WANTS_PLEASURE;
          } else if (category.type === "Extra") {
            targetRate = this.config.TARGET_RATES.EXTRA;
          }
          
          targetValues.push([targetRate]);
          
          // Calculate variance (actual % - target %)
          varianceFormulas.push([`=D${startRow + currentCategoryRow}-E${startRow + currentCategoryRow}`]);
          
          // Set background color
          backgroundColors.push([categoryBgColor, categoryBgColor, categoryBgColor, categoryBgColor, categoryBgColor, categoryBgColor]);
          
          // Add conditional formatting for the variance column
          conditionalFormatRules.push({
            row: startRow + currentCategoryRow,
            rule: SpreadsheetApp.newConditionalFormatRule()
              .whenFormulaSatisfied(`F${startRow + currentCategoryRow}>0`)
              .setBackground("#FFCDD2") // Light red if over budget
              .setRanges([this.analysisSheet.getRange(startRow + currentCategoryRow, 6)])
          });
          
          currentCategoryRow++;
        });
        
        // Add Total Expenses row
        categoryValues.push(["Total Expenses"]);
        typeValues.push(["All"]);
        amountValues.push([this.totals.expenses.value]);
        
        // Calculate percentage of income for total
        if (this.totals.income.value > 0) {
          percentFormulas.push([`=C${startRow + currentCategoryRow}/${this.totals.income.value}`]);
        } else {
          percentFormulas.push([0]);
        }
        
        targetValues.push([0.8]); // Target 80%
        varianceFormulas.push([`=D${startRow + currentCategoryRow}-E${startRow + currentCategoryRow}`]);
        
        // Set total row background
        backgroundColors.push([
          this.config.COLORS.UI.HEADER_BG, 
          this.config.COLORS.UI.HEADER_BG, 
          this.config.COLORS.UI.HEADER_BG, 
          this.config.COLORS.UI.HEADER_BG, 
          this.config.COLORS.UI.HEADER_BG, 
          this.config.COLORS.UI.HEADER_BG
        ]);
        
        currentCategoryRow++;
        
        // Write all data to the sheet in batches
        if (currentCategoryRow > 0) {
          // Set category names
          if (categoryValues.length > 0) {
            this.analysisSheet.getRange(startRow, 1, categoryValues.length, 1).setValues(categoryValues);
          }
          
          // Set types
          if (typeValues.length > 0) {
            this.analysisSheet.getRange(startRow, 2, typeValues.length, 1).setValues(typeValues);
          }
          
          // Set amounts
          if (amountValues.length > 0) {
            this.analysisSheet.getRange(startRow, 3, amountValues.length, 1).setValues(amountValues);
          }
          
          // Set percentage formulas
          if (percentFormulas.length > 0) {
            this.analysisSheet.getRange(startRow, 4, percentFormulas.length, 1).setFormulas(percentFormulas);
          }
          
          // Set target values
          if (targetValues.length > 0) {
            this.analysisSheet.getRange(startRow, 5, targetValues.length, 1).setValues(targetValues);
          }
          
          // Set variance formulas
          if (varianceFormulas.length > 0) {
            this.analysisSheet.getRange(startRow, 6, varianceFormulas.length, 1).setFormulas(varianceFormulas);
          }
          
          // Set backgrounds
          if (backgroundColors.length > 0) {
            this.analysisSheet.getRange(startRow, 1, backgroundColors.length, 6).setBackgrounds(backgroundColors);
          }
          
          // Format currency cells (amount column)
          utils.formatAsCurrency(this.analysisSheet.getRange(startRow, 3, currentCategoryRow, 1));
          
          // Format percentage cells (percentage columns)
          utils.formatAsPercentage(this.analysisSheet.getRange(startRow, 4, currentCategoryRow, 3));
          
          // Apply conditional formatting rules
          if (conditionalFormatRules.length > 0) {
            const rules = this.analysisSheet.getConditionalFormatRules();
            conditionalFormatRules.forEach(item => {
              rules.push(item.rule);
            });
            this.analysisSheet.setConditionalFormatRules(rules);
          }
          
          // Set font color for total row
          this.analysisSheet.getRange(startRow + currentCategoryRow - 1, 1, 1, 6)
            .setFontWeight("bold")
            .setFontColor(this.config.COLORS.UI.HEADER_FONT);
          
          // Add borders to the expense table
          this.analysisSheet.getRange(categoryStartRow, 1, currentCategoryRow, 6).setBorder(
            true, true, true, true, true, true, 
            "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID
          );
          
          // Add subtle shading to every other row for better readability
          for (let i = 0; i < currentCategoryRow - 1; i++) { // Skip total row
            if (i % 2 === 1) {
              this.analysisSheet.getRange(startRow + i, 1, 1, 6)
                .setBackground("#F8F8F8"); // Very light gray for alternate rows
            }
          }
        }
      }
      
      return startRow + currentCategoryRow;
    }

    /**
     * Creates and inserts expenditure charts (Pie chart for breakdown, Column chart for rates vs. targets)
     * into the analysis sheet.
     * @param {number} startRow - The 1-based row index where the charts should be positioned.
     * @return {void}
     * @private
     */
    createExpenditureCharts(startRow) {
      // Only create charts if we have expense categories
      if (this.data.expenseCategories.length === 0) return;
      
      // Find the rows containing category data in the analysis sheet
      const analysisData = this.analysisSheet.getDataRange().getValues();
      let categoryStartRow = -1;
      let categoryEndRow = -1;
      
      for (let i = 0; i < analysisData.length; i++) {
        if (analysisData[i][0] === "Category" && analysisData[i][1] === "Type") {
          categoryStartRow = i + 2; // Skip header row
        } else if (analysisData[i][0] === "Total Expenses" && analysisData[i][1] === "All") {
          categoryEndRow = i;
          break;
        }
      }
      
      if (categoryStartRow === -1 || categoryEndRow === -1) return;
      
      // Create a pie chart for expenditure breakdown
      const pieChartBuilder = this.analysisSheet.newChart();
      
      // Define chart data range (category name and amount)
      const pieDataRange = this.analysisSheet.getRange(categoryStartRow, 1, categoryEndRow - categoryStartRow, 3);
      
      pieChartBuilder
        .setChartType(Charts.ChartType.PIE)
        .addRange(pieDataRange)
        .setPosition(startRow, 1, 0, 0)
        .setOption('title', 'Expenditure Breakdown')
        .setOption('titleTextStyle', {
          color: this.config.COLORS.CHART.TITLE,
          fontSize: 16,
          bold: true
        })
        .setOption('pieSliceText', 'percentage')
        .setOption('pieHole', 0.4) // Create a donut chart for more modern look
        .setOption('legend', { 
          position: 'right',
          textStyle: {
            color: this.config.COLORS.CHART.TEXT,
            fontSize: 12
          }
        })
        .setOption('colors', this.config.COLORS.CHART.SERIES)
        .setOption('width', 450)
        .setOption('height', 300)
        .setOption('is3D', false)
        .setOption('pieSliceTextStyle', {
          color: '#FFFFFF',
          fontSize: 14,
          bold: true
        })
        .setOption('tooltip', { 
          showColorCode: true,
          textStyle: { fontSize: 12 }
        });
      
      // Add the pie chart to the sheet
      this.analysisSheet.insertChart(pieChartBuilder.build());
      
      // Create a column chart showing expense categories vs target
      const columnChartBuilder = this.analysisSheet.newChart();
      
      // Define data range for the column chart (category, actual %, target %)
      const columnDataRange = this.analysisSheet.getRange(categoryStartRow, 1, categoryEndRow - categoryStartRow, 5);
      
      columnChartBuilder
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(columnDataRange)
        .setPosition(startRow, 8, 0, 0)
        .setOption('title', 'Expense Rates vs Targets')
        .setOption('titleTextStyle', {
          color: this.config.COLORS.CHART.TITLE,
          fontSize: 16,
          bold: true
        })
        .setOption('legend', { 
          position: 'top',
          textStyle: {
            color: this.config.COLORS.CHART.TEXT,
            fontSize: 12
          }
        })
        .setOption('colors', [this.config.COLORS.UI.EXPENSE_FONT, this.config.COLORS.UI.INCOME_FONT]) // Red for actual, green for target
        .setOption('width', 450)
        .setOption('height', 300)
        .setOption('hAxis', {
          title: 'Category',
          titleTextStyle: {color: this.config.COLORS.CHART.TEXT},
          textStyle: {color: this.config.COLORS.CHART.TEXT, fontSize: 10}
        })
        .setOption('vAxis', {
          title: 'Rate (% of Income)',
          titleTextStyle: {color: this.config.COLORS.CHART.TEXT},
          textStyle: {color: this.config.COLORS.CHART.TEXT},
          format: 'percent'
        })
        .setOption('bar', {groupWidth: '75%'})
        .setOption('isStacked', false);
      
      // Add the column chart to the sheet
      this.analysisSheet.insertChart(columnChartBuilder.build());
    }

    /**
     * Placeholder method for suggesting savings opportunities. Currently shows an alert.
     * @return {void}
     * @public
     */
    suggestSavingsOpportunities() {
      // TODO: Implement savings opportunities suggestion
      SpreadsheetApp.getUi().alert('Savings Opportunities - Coming Soon!');
    }

    /**
     * Placeholder method for detecting spending anomalies. Currently shows an alert.
     * @return {void}
     * @public
     */
    detectSpendingAnomalies() {
      // TODO: Implement spending anomaly detection
      SpreadsheetApp.getUi().alert('Spending Anomalies Detection - Coming Soon!');
    }

    /**
     * Placeholder method for analyzing fixed vs. variable expenses. Currently shows an alert.
     * @return {void}
     * @public
     */
    analyzeFixedVsVariableExpenses() {
      // TODO: Implement fixed vs variable expenses analysis
      SpreadsheetApp.getUi().alert('Fixed vs Variable Expenses Analysis - Coming Soon!');
    }

    /**
     * Placeholder method for generating a cash flow forecast. Currently shows an alert.
     * @return {void}
     * @public
     */
    generateCashFlowForecast() {
      // TODO: Implement cash flow forecast
      SpreadsheetApp.getUi().alert('Cash Flow Forecast - Coming Soon!');
    }
  }
  
  // ============================================================================
  // PUBLIC API
  // ============================================================================
  
  return {
    /**
     * Creates an instance of the internal `FinancialAnalysisService` class, initializes it,
     * and runs the analysis workflow. This is the primary entry point for generating the analysis.
     * Provides UI feedback during the process.
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet object.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} overviewSheet - The sheet object containing the generated financial overview data.
     * @return {FinancialAnalysisService} An initialized instance of the internal `FinancialAnalysisService` class,
     *         which contains the analysis results and methods for further interaction (though currently limited).
     * @throws {Error} Re-throws any error encountered during analysis after logging and notifying the user.
     * @public
     * @example
     * // Typically called internally by other services or controllers
     * const analysisService = FinancialPlanner.FinancialAnalysisService.analyze(ss, overviewSheet);
     */
    analyze: function(spreadsheet, overviewSheet) {
      try {
        uiService.showLoadingSpinner("Analyzing financial data...");
        
        const analysisConfig = {
          ...config.get(),
          // Add any additional config needed
          TARGET_RATES: {
            ...config.getSection('TARGET_RATES'),
            WANTS_PLEASURE: config.getSection('TARGET_RATES').WANTS // Map WANTS to WANTS_PLEASURE for compatibility
          }
        };
        
        const service = new FinancialAnalysisService(
          spreadsheet,
          overviewSheet,
          analysisConfig
        );
        
        service.initialize();
        service.analyze();
        
        uiService.hideLoadingSpinner();
        
        return service;
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Error analyzing financial data");
        throw error;
      }
    },
    
    /**
     * Public method to trigger the generation and display of the key metrics analysis.
     * It ensures the overview sheet exists, calls the internal `analyze` method,
     * activates the analysis sheet, and shows a success notification.
     * @return {void}
     * @public
     * @example
     * // Called from a menu item or controller:
     * FinancialPlanner.FinancialAnalysisService.showKeyMetrics();
     */
    showKeyMetrics: function() {
      try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const overviewSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').OVERVIEW);
        
        if (!overviewSheet) {
          uiService.showErrorNotification("Error", 
            "Overview sheet not found. Please generate the financial overview first.");
          return;
        }
        
        // Analyze the data by calling the module's analyze function
        // which correctly instantiates the service
        FinancialPlanner.FinancialAnalysisService.analyze(spreadsheet, overviewSheet);
        
        // Activate the Analysis sheet to show it to the user
        const analysisSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').ANALYSIS);
        analysisSheet.activate();
        
        uiService.showSuccessNotification("Key metrics have been generated in the Analysis sheet.");
      } catch (error) {
        errorService.handle(error, "Failed to generate key metrics");
      }
    },
    
    /**
     * Public method to trigger the suggestion of savings opportunities.
     * Ensures the overview sheet exists, runs the analysis, and calls the internal placeholder method.
     * @return {void}
     * @public
     * @example
     * // Called from a menu item or controller:
     * FinancialPlanner.FinancialAnalysisService.suggestSavingsOpportunities();
     */
    suggestSavingsOpportunities: function() {
      try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const overviewSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').OVERVIEW);
        
        if (!overviewSheet) {
          uiService.showErrorNotification("Error", 
            "Overview sheet not found. Please generate the financial overview first.");
          return;
        }
        
        // Create and use the service
        const service = this.analyze(spreadsheet, overviewSheet);
        service.suggestSavingsOpportunities();
      } catch (error) {
        errorService.handle(error, "Failed to suggest savings opportunities");
      }
    },
    
    /**
     * Public method to trigger the detection of spending anomalies.
     * Ensures the overview sheet exists, runs the analysis, and calls the internal placeholder method.
     * @return {void}
     * @public
     * @example
     * // Called from a menu item or controller:
     * FinancialPlanner.FinancialAnalysisService.detectSpendingAnomalies();
     */
    detectSpendingAnomalies: function() {
      try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const overviewSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').OVERVIEW);
        
        if (!overviewSheet) {
          uiService.showErrorNotification("Error", 
            "Overview sheet not found. Please generate the financial overview first.");
          return;
        }
        
        // Create and use the service
        const service = this.analyze(spreadsheet, overviewSheet);
        service.detectSpendingAnomalies();
      } catch (error) {
        errorService.handle(error, "Failed to detect spending anomalies");
      }
    }
  };
})(FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.ErrorService, FinancialPlanner.Config);

// ============================================================================
// BACKWARD COMPATIBILITY LAYER
// ============================================================================

/**
 * Shows the key metrics section in the Analysis sheet.
 * Maintained for backward compatibility with older triggers or direct calls.
 * Delegates to `FinancialPlanner.FinancialAnalysisService.showKeyMetrics()`.
 * @return {void}
 * @global
 */
function showKeyMetrics() {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.FinancialAnalysisService && FinancialPlanner.FinancialAnalysisService.showKeyMetrics) {
    FinancialPlanner.FinancialAnalysisService.showKeyMetrics();
  } else {
     Logger.log("Global showKeyMetrics: FinancialPlanner.FinancialAnalysisService not available.");
     // Optionally show an error to the user if appropriate for a direct call scenario
     // SpreadsheetApp.getUi().alert("Error: Financial Analysis module not loaded.");
  }
}
