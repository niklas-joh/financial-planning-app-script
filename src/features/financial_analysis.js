/**
 * Financial Planning Tools - Financial Analysis Service
 * 
 * This module provides analytics functionality for financial data through a dedicated
 * FinancialAnalysisService class. It creates a separate Analysis sheet with key metrics,
 * expense category analysis, and visualizations.
 */

/**
 * FinancialAnalysisService class for handling all financial analytics functionality
 * @class
 */
class FinancialAnalysisService {
  /**
   * Creates a new FinancialAnalysisService instance
   * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The active spreadsheet
   * @param {SpreadsheetApp.Sheet} overviewSheet - The overview sheet containing financial data
   * @param {Object} config - Configuration object with settings and constants
   * @constructor
   */
  constructor(spreadsheet, overviewSheet, config) {
    /**
     * The active spreadsheet
     * @type {SpreadsheetApp.Spreadsheet}
     * @private
     */
    this.spreadsheet = spreadsheet;
    
    /**
     * The overview sheet containing financial data
     * @type {SpreadsheetApp.Sheet}
     * @private
     */
    this.overviewSheet = overviewSheet;
    
    /**
     * Configuration object with settings and constants
     * @type {Object}
     * @private
     */
    this.config = config;
    
    /**
     * The analysis sheet where analytics will be displayed
     * @type {SpreadsheetApp.Sheet}
     * @private
     */
    this.analysisSheet = getOrCreateSheet(spreadsheet, config.SHEETS.ANALYSIS);
    
    /**
     * Extracted financial data for analysis
     * @type {Object|null}
     * @private
     */
    this.data = null;
    
    /**
     * Calculated totals for different financial categories
     * @type {Object|null}
     * @private
     */
    this.totals = null;
  }

  /**
   * Initializes the analysis service by extracting data and setting up the analysis sheet
   * @public
   */
  initialize() {
    // Extract data from the overview sheet for analysis
    this.extractDataFromOverview();
    
    // Clear and set up the analysis sheet
    this.setupAnalysisSheet();
  }

  /**
   * Performs all analysis functions to generate the complete analysis
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
   * Extracts necessary data from the overview sheet for analysis
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
   * Sets up the analysis sheet with header and formatting
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
   * Adds key metrics section to the analysis sheet
   * @param {Number} startRow - The row to start adding key metrics
   * @returns {Number} The next row index after adding the section
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
    
    // Add metrics rows
    // 1. Savings Rate
    if (this.totals.income.row > 0 && this.totals.savings.row > 0) {
      this.analysisSheet.getRange(startRow, 1).setValue("Savings Rate");
      this.analysisSheet.getRange(startRow, 2).setFormula(
        `=-${this.totals.savings.value}/${this.totals.income.value}`
      );
      this.analysisSheet.getRange(startRow, 3).setValue(this.config.TARGET_RATES.DEFAULT); // Use config target rate
      this.analysisSheet.getRange(startRow, 4).setValue(
        "Positive % indicates saving money, negative % indicates withdrawing from savings"
      );
      this.analysisSheet.getRange(startRow, 1, 1, 4).setBackground(this.config.COLORS.UI.METRICS_BG);
      
      // Format as percentage
      formatAsPercentage(this.analysisSheet.getRange(startRow, 2, 1, 2));
      
      // Add conditional formatting (green if meeting target, red if not)
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(this.analysisSheet.getRange(startRow, 3).getValue())
        .setBackground("#FFCDD2") // Light red if below target
        .setRanges([this.analysisSheet.getRange(startRow, 2)])
        .build();
      
      const rules = this.analysisSheet.getConditionalFormatRules();
      rules.push(rule);
      this.analysisSheet.setConditionalFormatRules(rules);
      
      startRow++;
    }
    
    // 2. Expenses/Income Ratio
    if (this.totals.income.row > 0 && this.totals.expenses.row > 0) {
      this.analysisSheet.getRange(startRow, 1).setValue("Expenses/Income Ratio");
      this.analysisSheet.getRange(startRow, 2).setFormula(
        `=-${this.totals.expenses.value}/${this.totals.income.value}`
      );
      this.analysisSheet.getRange(startRow, 3).setValue(this.config.TARGET_RATES.DEFAULT * -1); // Use config target rate with negative sign
      this.analysisSheet.getRange(startRow, 4).setValue(
        "Negative % indicates spending, lower absolute value is better"
      );
      this.analysisSheet.getRange(startRow, 1, 1, 4).setBackground(startRow % 2 === 0 ? "#F5F5F5" : this.config.COLORS.UI.METRICS_BG);
      
      // Format as percentage
      formatAsPercentage(this.analysisSheet.getRange(startRow, 2, 1, 2));
      
      // Add conditional formatting (green if meeting target, red if not)
      // Since we reversed the sign, we need to adjust the conditional formatting
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(this.analysisSheet.getRange(startRow, 3).getValue())
        .setBackground("#FFCDD2") // Light red if below target (more negative)
        .setRanges([this.analysisSheet.getRange(startRow, 2)])
        .build();
      
      const rules = this.analysisSheet.getConditionalFormatRules();
      rules.push(rule);
      this.analysisSheet.setConditionalFormatRules(rules);
      
      startRow++;
    }
    
    // 3. Monthly Net Cash Flow
    if (this.totals.income.row > 0 && this.totals.expenses.row > 0) {
      this.analysisSheet.getRange(startRow, 1).setValue("Monthly Net Cash Flow");
      this.analysisSheet.getRange(startRow, 2).setFormula(
        `=${this.totals.income.value}-${this.totals.expenses.value}`
      );
      this.analysisSheet.getRange(startRow, 3).setValue(0); // Target is positive cash flow
      this.analysisSheet.getRange(startRow, 4).setValue(
        "Positive value means you're earning more than spending, negative means you're spending more than earning"
      );
      this.analysisSheet.getRange(startRow, 1, 1, 4).setBackground(startRow % 2 === 0 ? "#F5F5F5" : this.config.COLORS.UI.METRICS_BG);
      
      // Format as currency
      formatAsCurrency(this.analysisSheet.getRange(startRow, 2, 1, 2));
      
      // Add conditional formatting (green if positive, red if negative)
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground("#FFCDD2") // Light red if negative
        .setRanges([this.analysisSheet.getRange(startRow, 2)])
        .build();
      
      const rules = this.analysisSheet.getConditionalFormatRules();
      rules.push(rule);
      this.analysisSheet.setConditionalFormatRules(rules);
      
      startRow++;
    }
    
    // 4. Essentials Rate (new)
    if (this.totals.income.row > 0 && this.totals.essentials.row > 0) {
      this.analysisSheet.getRange(startRow, 1).setValue("Essentials Rate");
      this.analysisSheet.getRange(startRow, 2).setFormula(
        `=${this.totals.essentials.value}/${this.totals.income.value}`
      );
      this.analysisSheet.getRange(startRow, 3).setValue(this.config.TARGET_RATES.ESSENTIALS);
      this.analysisSheet.getRange(startRow, 4).setValue(
        "Percentage of income spent on essential expenses (lower is better)"
      );
      this.analysisSheet.getRange(startRow, 1, 1, 4).setBackground(startRow % 2 === 0 ? "#F5F5F5" : this.config.COLORS.UI.METRICS_BG);
      
      // Format as percentage
      formatAsPercentage(this.analysisSheet.getRange(startRow, 2, 1, 2));
      
      // Add conditional formatting (red if exceeding target)
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(this.analysisSheet.getRange(startRow, 3).getValue())
        .setBackground("#FFCDD2") // Light red if above target
        .setRanges([this.analysisSheet.getRange(startRow, 2)])
        .build();
      
      const rules = this.analysisSheet.getConditionalFormatRules();
      rules.push(rule);
      this.analysisSheet.setConditionalFormatRules(rules);
      
      startRow++;
    }
    
    // 5. Wants/Pleasure Rate (new)
    if (this.totals.income.row > 0 && this.totals.wantsPleasure.row > 0) {
      this.analysisSheet.getRange(startRow, 1).setValue("Wants/Pleasure Rate");
      this.analysisSheet.getRange(startRow, 2).setFormula(
        `=${this.totals.wantsPleasure.value}/${this.totals.income.value}`
      );
      this.analysisSheet.getRange(startRow, 3).setValue(this.config.TARGET_RATES.WANTS_PLEASURE);
      this.analysisSheet.getRange(startRow, 4).setValue(
        "Percentage of income spent on wants and pleasure (discretionary spending)"
      );
      this.analysisSheet.getRange(startRow, 1, 1, 4).setBackground(startRow % 2 === 0 ? "#F5F5F5" : this.config.COLORS.UI.METRICS_BG);
      
      // Format as percentage
      formatAsPercentage(this.analysisSheet.getRange(startRow, 2, 1, 2));
      
      // Add conditional formatting (red if exceeding target)
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(this.analysisSheet.getRange(startRow, 3).getValue())
        .setBackground("#FFCDD2") // Light red if above target
        .setRanges([this.analysisSheet.getRange(startRow, 2)])
        .build();
      
      const rules = this.analysisSheet.getConditionalFormatRules();
      rules.push(rule);
      this.analysisSheet.setConditionalFormatRules(rules);
      
      startRow++;
    }
    
    // 6. Extra Expenses Rate (new)
    if (this.totals.income.row > 0 && this.totals.extra.row > 0) {
      this.analysisSheet.getRange(startRow, 1).setValue("Extra Expenses Rate");
      this.analysisSheet.getRange(startRow, 2).setFormula(
        `=${this.totals.extra.value}/${this.totals.income.value}`
      );
      this.analysisSheet.getRange(startRow, 3).setValue(this.config.TARGET_RATES.EXTRA);
      this.analysisSheet.getRange(startRow, 4).setValue(
        "Percentage of income spent on extra/miscellaneous expenses"
      );
      this.analysisSheet.getRange(startRow, 1, 1, 4).setBackground(startRow % 2 === 0 ? "#F5F5F5" : this.config.COLORS.UI.METRICS_BG);
      
      // Format as percentage
      formatAsPercentage(this.analysisSheet.getRange(startRow, 2, 1, 2));
      
      // Add conditional formatting (red if exceeding target)
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(this.analysisSheet.getRange(startRow, 3).getValue())
        .setBackground("#FFCDD2") // Light red if above target
        .setRanges([this.analysisSheet.getRange(startRow, 2)])
        .build();
      
      const rules = this.analysisSheet.getConditionalFormatRules();
      rules.push(rule);
      this.analysisSheet.setConditionalFormatRules(rules);
      
      startRow++;
    }
    
    // Add a border around the metrics table
    const metricsRowCount = startRow - 3; // Calculate the number of metrics rows
    this.analysisSheet.getRange(startRow - metricsRowCount, 1, metricsRowCount, 4).setBorder(
      true, true, true, true, true, true, 
      "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID
    );
    
    return startRow;
  }

  /**
   * Adds expense categories section to the analysis sheet
   * @param {Number} startRow - The row to start adding expense categories
   * @returns {Number} The next row index after adding the section
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
    this.analysisSheet.getRange(startRow, 1).setValue("Category");
    this.analysisSheet.getRange(startRow, 2).setValue("Type");
    this.analysisSheet.getRange(startRow, 3).setValue("Amount");
    this.analysisSheet.getRange(startRow, 4).setValue("% of Income");
    this.analysisSheet.getRange(startRow, 5).setValue("Target %");
    this.analysisSheet.getRange(startRow, 6).setValue("Variance");
    
    this.analysisSheet.getRange(startRow, 1, 1, 6)
      .setBackground("#F5F5F5")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    startRow++;
    
    // Add rows for each expense category
    if (this.data.expenseCategories.length > 0) {
      // Sort categories by amount (descending)
      const sortedCategories = [...this.data.expenseCategories].sort((a, b) => b.amount - a.amount);
      
      // Add a row for each category
      sortedCategories.forEach((category, index) => {
        // Skip subcategories for simplicity
        if (category.subcategory) return;
        
        this.analysisSheet.getRange(startRow, 1).setValue(category.category);
        this.analysisSheet.getRange(startRow, 2).setValue(category.type);
        this.analysisSheet.getRange(startRow, 3).setValue(category.amount);
        
        // Calculate percentage of income
        if (this.totals.income.value > 0) {
          this.analysisSheet.getRange(startRow, 4).setFormula(`=C${startRow}/${this.totals.income.value}`);
        } else {
          this.analysisSheet.getRange(startRow, 4).setValue(0);
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
        
        this.analysisSheet.getRange(startRow, 5).setValue(targetRate);
        
        // Calculate variance (actual % - target %)
        this.analysisSheet.getRange(startRow, 6).setFormula(`=D${startRow}-E${startRow}`);
        
        // Apply alternating row colors
        this.analysisSheet.getRange(startRow, 1, 1, 6).setBackground(index % 2 === 0 ? "#F5F5F5" : this.config.COLORS.UI.METRICS_BG);
        
        // Format cells
        formatAsCurrency(this.analysisSheet.getRange(startRow, 3)); // Amount column as currency
        formatAsPercentage(this.analysisSheet.getRange(startRow, 4, 1, 3)); // Percentage columns
        
        // Add conditional formatting for the variance column
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground("#FFCDD2") // Light red if over budget
          .setRanges([this.analysisSheet.getRange(startRow, 6)])
          .build();
        
        const rules = this.analysisSheet.getConditionalFormatRules();
        rules.push(rule);
        this.analysisSheet.setConditionalFormatRules(rules);
        
        startRow++;
      });
      
      // Add Total Expenses row with distinct formatting
      this.analysisSheet.getRange(startRow, 1).setValue("Total Expenses");
      this.analysisSheet.getRange(startRow, 2).setValue("All");
      this.analysisSheet.getRange(startRow, 3).setValue(this.totals.expenses.value);
      
      // Calculate percentage of income
      if (this.totals.income.value > 0) {
        this.analysisSheet.getRange(startRow, 4).setFormula(`=C${startRow}/${this.totals.income.value}`);
      } else {
        this.analysisSheet.getRange(startRow, 4).setValue(0);
      }
      
      this.analysisSheet.getRange(startRow, 5).setValue(0.8); // Target 80%
      this.analysisSheet.getRange(startRow, 6).setFormula(`=D${startRow}-E${startRow}`);
      
      // Format the total row
      this.analysisSheet.getRange(startRow, 1, 1, 6)
        .setBackground(this.config.COLORS.UI.HEADER_BG)
        .setFontWeight("bold")
        .setFontColor(this.config.COLORS.UI.HEADER_FONT);
      
      // Format cells
      formatAsCurrency(this.analysisSheet.getRange(startRow, 3));
      formatAsPercentage(this.analysisSheet.getRange(startRow, 4, 1, 3));
      
      startRow++;
    }
    
    // Add borders to the expense table
    const tableStartRow = startRow - this.data.expenseCategories.length - 1;
    this.analysisSheet.getRange(tableStartRow, 1, startRow - tableStartRow, 6).setBorder(
      true, true, true, true, true, true, 
      "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID
    );
    
    return startRow;
  }

  /**
   * Creates expenditure charts on the analysis sheet
   * @param {Number} startRow - The row to start adding charts
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
   * Suggests savings opportunities based on spending patterns
   * @public
   */
  suggestSavingsOpportunities() {
    // TODO: Implement savings opportunities suggestion
    SpreadsheetApp.getUi().alert('Savings Opportunities - Coming Soon!');
  }

  /**
   * Detects spending anomalies in transaction data
   * @public
   */
  detectSpendingAnomalies() {
    // TODO: Implement spending anomaly detection
    SpreadsheetApp.getUi().alert('Spending Anomalies Detection - Coming Soon!');
  }

  /**
   * Analyzes fixed vs variable expenses
   * @public
   */
  analyzeFixedVsVariableExpenses() {
    // TODO: Implement fixed vs variable expenses analysis
    SpreadsheetApp.getUi().alert('Fixed vs Variable Expenses Analysis - Coming Soon!');
  }

  /**
   * Generates a cash flow forecast based on historical data
   * @public
   */
  generateCashFlowForecast() {
    // TODO: Implement cash flow forecast
    SpreadsheetApp.getUi().alert('Cash Flow Forecast - Coming Soon!');
  }
}

// Export the class for use in other modules
// Note: In Google Apps Script, functions are automatically global
