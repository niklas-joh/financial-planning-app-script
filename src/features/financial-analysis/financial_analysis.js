/**
 * @fileoverview Financial Analysis Service for the Financial Planning Tools.
 * This module provides functionality to generate a financial analysis sheet,
 * including key metrics, expense breakdowns, and charts, based on data
 * from an overview sheet. It utilizes various other services for data extraction,
 * sheet building, and calculations.
 * Version: 3.0.0
 * @module features/financial-analysis
 */

/**
 * @namespace FinancialPlanner.FinancialAnalysisService
 * @description Service for performing and displaying financial analysis.
 * This IIFE sets up the service with its dependencies.
 * @param {UtilsModule} utils - Instance of the Utils module.
 * @param {UIServiceModule} uiService - Instance of the UI Service module.
 * @param {ErrorServiceModule} errorService - Instance of the Error Service module.
 * @param {ConfigModule} config - Instance of the Config module.
 * @param {SheetBuilderModule} sheetBuilder - Instance of the Sheet Builder module (factory).
 * @param {MetricsCalculatorModule} metricsCalculator - Instance of the Metrics Calculator module.
 * @param {FormulaBuilderModule} formulaBuilder - Instance of the Formula Builder module.
 */
FinancialPlanner.FinancialAnalysisService = (function(
  utils, uiService, errorService, config, 
  sheetBuilder, metricsCalculator, formulaBuilder
) {
  
  /**
   * @classdesc Extracts structured financial metrics and category data from a financial overview sheet.
   * @class DataExtractor
   * @private
   */
  class DataExtractor {
    /**
     * Creates an instance of DataExtractor.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} overviewSheet - The sheet object for the financial overview.
     */
    constructor(overviewSheet) {
      /** @type {GoogleAppsScript.Spreadsheet.Sheet} */
      this.sheet = overviewSheet;
      /** @type {Array<Array<*>>} */
      this.data = this.sheet.getDataRange().getValues();
      /** @type {string} Sheet name quoted for use in formulas. */
      this.overviewSheetName = `'${config.getSection('SHEETS').OVERVIEW}'`;
    }
    
    /**
     * Extracts key financial metrics from the overview sheet data.
     * @returns {{
     *   income: ReturnType<DataExtractor['findTotalRow']>,
     *   expenses: ReturnType<DataExtractor['findExpenseRows']>,
     *   savings: ReturnType<DataExtractor['findTotalRow']>,
     *   categories: ReturnType<DataExtractor['extractCategories']>,
     *   essentials: ReturnType<DataExtractor['findTotalRow']>,
     *   wantsPleasure: ReturnType<DataExtractor['findTotalRow']>,
     *   extra: ReturnType<DataExtractor['findTotalRow']>
     * }} An object containing extracted metrics.
     * @memberof DataExtractor
     */
    extractMetrics() {
      const metrics = {
        income: this.findTotalRow('Total Income'),
        expenses: this.findExpenseRows(),
        savings: this.findTotalRow('Total Savings'),
        categories: this.extractCategories()
      };
      
      // Find specific expense types
      metrics.essentials = this.findTotalRow('Essentials');
      metrics.wantsPleasure = this.findTotalRow('Wants/Pleasure');
      metrics.extra = this.findTotalRow('Extra');
      
      return metrics;
    }
    
    /**
     * Finds a row by its label in the first column and extracts total, average, and cell references.
     * @param {string} label - The label to search for in the first column.
     * @returns {{row: number, total: number, average: number, totalRef: string, averageRef: string, monthlyValuesRangeRef: string}|null}
     *   An object with row data or null if label not found.
     * @memberof DataExtractor
     */
    findTotalRow(label) {
      const totalColLetter = utils.columnToLetter(17);
      const averageColLetter = utils.columnToLetter(18);
      const monthlyStartCol = utils.columnToLetter(5);
      const monthlyEndCol = utils.columnToLetter(16);
      
      for (let i = 0; i < this.data.length; i++) {
        if (this.data[i][0] === label) {
          const rowNum = i + 1;
          return {
            row: rowNum,
            total: this.data[i][16],
            average: this.data[i][17],
            totalRef: `${this.overviewSheetName}!${totalColLetter}${rowNum}`,
            averageRef: `${this.overviewSheetName}!${averageColLetter}${rowNum}`,
            monthlyValuesRangeRef: `${this.overviewSheetName}!${monthlyStartCol}${rowNum}:${monthlyEndCol}${rowNum}`
          };
        }
      }
      return null;
    }
    
    /**
     * Finds the 'Total Expenses' row and extracts its data.
     * @returns {{row: number, total: number, average: number, totalRef: string, averageRef: string, monthlyValuesRangeRef: string}}
     *   An object with total expenses data. Returns a default object if not found.
     * @memberof DataExtractor
     */
    findExpenseRows() {
      let totalExpenses = { row: -1, total: 0, average: 0, totalRef: '', averageRef: '', monthlyValuesRangeRef: '' };
      
      for (let i = 0; i < this.data.length; i++) {
        if (this.data[i][0] === 'Total Expenses') {
          const rowNum = i + 1;
          totalExpenses = {
            row: rowNum,
            total: this.data[i][16],
            average: this.data[i][17],
            totalRef: `${this.overviewSheetName}!${utils.columnToLetter(17)}${rowNum}`,
            averageRef: `${this.overviewSheetName}!${utils.columnToLetter(18)}${rowNum}`,
            monthlyValuesRangeRef: `${this.overviewSheetName}!${utils.columnToLetter(5)}${rowNum}:${utils.columnToLetter(16)}${rowNum}`
          };
          break;
        }
      }
      
      return totalExpenses;
    }
    
    /**
     * Extracts and categorizes income, expense, and savings items from the overview data.
     * @returns {{
     *   income: Array<{type: string, category: string, subcategory: string, amount: number, row: number}>,
     *   expenses: Array<{type: string, category: string, subcategory: string, amount: number, row: number}>,
     *   savings: Array<{type: string, category: string, subcategory: string, amount: number, row: number}>
     * }} An object containing arrays of categorized financial items.
     * @memberof DataExtractor
     */
    extractCategories() {
      const categories = {
        income: [],
        expenses: [],
        savings: []
      };
      
      // Iterate through all rows in the overview sheet data
      for (let i = 0; i < this.data.length; i++) {
        const row = this.data[i];
        const type = row[0];
        const category = row[1];
        const subcategory = row[2];
        const average = row[17];
        
        if (!category) continue;
        
        const categoryData = {
          type: type,
          category: category,
          subcategory: subcategory || '',
          amount: average,
          row: i + 1
        };
        
        if (type === 'Income') {
          categories.income.push(categoryData);
        } else if (type === 'Savings') {
          categories.savings.push(categoryData);
        } else if (config.getSection('EXPENSE_TYPES').includes(type)) { // Check if the type is a known expense type from config
          categories.expenses.push(categoryData);
        }
      }
      
      return categories;
    }
  }
  
  /**
   * @classdesc Builds metric cards on a sheet using a SheetBuilder instance.
   * @class MetricCardBuilder
   * @private
   */
  class MetricCardBuilder {
    /**
     * Creates an instance of MetricCardBuilder.
     * @param {SheetBuilder} builder - An instance of the SheetBuilder service.
     * @param {ConfigModule} config - An instance of the Config module.
     * @param {MetricsCalculatorModule} calculator - An instance of the MetricsCalculator module.
     */
    constructor(builder, config, calculator) {
      /** @type {SheetBuilder} */
      this.builder = builder;
      /** @type {ConfigModule} */
      this.config = config;
      /** @type {MetricsCalculatorModule} */
      this.calculator = calculator;
    }
    
    /**
     * Creates a grid of metric card pairs.
     * @param {Array<{cashFlow: object, rate: object}>} metrics - An array of metric pair configurations.
     * @param {number} startRow - The starting row on the sheet to build the card grid.
     * @memberof MetricCardBuilder
     */
    createCardGrid(metrics, startRow) {
      this.builder.setCurrentRow(startRow);
      
      // Create two columns of cards
      metrics.forEach((pair, index) => {
        this.createCardPair(pair.cashFlow, pair.rate);
        // Add spacing between card pairs
        this.builder.addBlankRow(15);
      });
    }
    
    /**
     * Creates a pair of metric cards (one cash flow, one rate) side-by-side.
     * @param {object} cashFlowMetric - Configuration for the cash flow metric card.
     * @param {object} rateMetric - Configuration for the rate metric card.
     * @memberof MetricCardBuilder
     */
    createCardPair(cashFlowMetric, rateMetric) {
      const currentRow = this.builder.getCurrentRow();
      
      // Create cash flow card (columns B-C)
      this.createCard(cashFlowMetric, currentRow, 2);
      
      // Create rate card (columns E-F)
      this.createCard(rateMetric, currentRow, 5);
      
      this.builder.setCurrentRow(currentRow + 5);
    }
    
    /**
     * Creates a single metric card with a specific layout and formatting.
     * @param {{
     *   name: string,
     *   avgFormula?: string,
     *   totalFormula?: string,
     *   targetValue?: number,
     *   sparklinePlaceholderText?: string,
     *   valueType: 'currency'|'percentage',
     *   description?: string,
     *   avgLabel?: string,
     *   totalLabel?: string
     * }} metric - Configuration object for the metric card.
     * @param {number} startRow - The starting row for this card.
     * @param {number} startColumn - The starting column for this card.
     * @memberof MetricCardBuilder
     */
    createCard(metric, startRow, startColumn) {
      const sheet = this.builder.sheet;
      
      // Define standard row heights for consistent card appearance
      sheet.setRowHeight(startRow, 25);      // Card Header row
      sheet.setRowHeight(startRow + 1, 35);  // Main values (monthly/annual) row
      sheet.setRowHeight(startRow + 2, 30);  // Sparkline or trend indicator row
      sheet.setRowHeight(startRow + 3, 20);  // Labels for values row
      sheet.setRowHeight(startRow + 4, 35);  // Description text row
      
      // Card Header: Merged cells with title and specific styling
      sheet.getRange(startRow, startColumn, 1, 2).merge()
        .setValue(metric.name)
        .setBackground('#E2EFDA') // Light green background for header
        .setFontWeight('bold')
        .setFontColor('black')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Monthly Value Display: Formula-driven, large font, green color
      const monthlyValueCell = sheet.getRange(startRow + 1, startColumn);
      monthlyValueCell.setFormula(metric.avgFormula)
        .setFontSize(18)
        .setFontColor('#008000') // Green color for positive financial figures
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Annual/Target Value Display: Formula or direct value, similar styling
      const annualValueCell = sheet.getRange(startRow + 1, startColumn + 1);
      if (metric.totalFormula) {
        annualValueCell.setFormula(metric.totalFormula);
      } else if (metric.targetValue !== undefined) {
        annualValueCell.setValue(metric.targetValue);
      }
      annualValueCell.setFontSize(18)
        .setFontColor('#008000') // Green color
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Apply Number Formatting based on metric type (currency or percentage)
      if (metric.valueType === 'currency') {
        const currencyFormat = this.config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT;
        monthlyValueCell.setNumberFormat(currencyFormat);
        if (metric.totalFormula || typeof metric.targetValue === 'number') annualValueCell.setNumberFormat(currencyFormat);
      } else if (metric.valueType === 'percentage') {
        monthlyValueCell.setNumberFormat('0.00%');
        if (metric.totalFormula || typeof metric.targetValue === 'number') { // Check if targetValue is a number for formatting
          annualValueCell.setNumberFormat('0.00%');
        }
      }
      
      // Sparkline Placeholder: Merged cells with placeholder text
      sheet.getRange(startRow + 2, startColumn, 1, 2).merge()
        .setValue(metric.sparklinePlaceholderText || `[Sparkline: ${metric.name}]`)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setFontStyle('italic')
        .setFontColor('#AAAAAA'); // Grey color for placeholder
      
      // Value Labels: Small text for "Monthly Avg.", "Target", etc.
      sheet.getRange(startRow + 3, startColumn)
        .setValue(metric.avgLabel || 'Monthly Avg.')
        .setFontSize(9)
        .setFontColor('#808080') // Grey color for labels
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      sheet.getRange(startRow + 3, startColumn + 1)
        .setValue(metric.totalLabel || (metric.targetValue !== undefined ? 'Target' : 'Annual Total'))
        .setFontSize(9)
        .setFontColor('#808080') // Grey color for labels
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Description Area: Merged cells for a brief explanation of the metric
      sheet.getRange(startRow + 4, startColumn, 1, 2).merge()
        .setValue(metric.description || '')
        .setFontSize(9)
        .setFontColor('#595959') // Darker grey for description text
        .setHorizontalAlignment('center')
        .setVerticalAlignment('top')
        .setWrap(true);
      
      // Card Border: Apply a thin solid border around the entire card
      sheet.getRange(startRow, startColumn, 5, 2)
        .setBorder(true, true, true, true, true, true, '#A9D18E', SpreadsheetApp.BorderStyle.SOLID_THIN); // Light green border
    }
  }
  
  /**
   * @classdesc Builds charts (e.g., pie, column) on a sheet.
   * @class ChartBuilder
   * @private
   */
  class ChartBuilder {
    /**
     * Creates an instance of ChartBuilder.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where charts will be inserted.
     * @param {ConfigModule} config - An instance of the Config module for chart styling.
     */
    constructor(sheet, config) {
      /** @type {GoogleAppsScript.Spreadsheet.Sheet} */
      this.sheet = sheet;
      /** @type {ConfigModule} */
      this.config = config;
    }
    
    /**
     * Creates and inserts expense breakdown charts (pie and column) onto the sheet.
     * @param {Array<object>} categoryData - Data for expense categories (not directly used, reads from sheet).
     * @param {number} startRow - The target row for positioning the top-left corner of the charts.
     * @memberof ChartBuilder
     */
    createExpenseCharts(categoryData, startRow) {
      // Dynamically find the data range for expense categories in the analysis sheet.
      // This relies on specific header text ("Category", "Type") and a total row ("Total Expenses", "All")
      // to define the boundaries of the category data used for charting.
      const analysisData = this.sheet.getDataRange().getValues();
      let categoryStartRow = -1; // 1-based row index in sheet, after header
      let categoryEndRow = -1;   // 1-based row index in sheet, for last category item

      for (let i = 0; i < analysisData.length; i++) {
        if (analysisData[i][0] === "Category" && analysisData[i][1] === "Type") { // Header row for category table
          categoryStartRow = i + 2; // Data starts two rows below this header (i is 0-based, sheet is 1-based)
        } else if (analysisData[i][0] === "Total Expenses" && analysisData[i][1] === "All") { // Row after last category
          categoryEndRow = i; // This is the row index (0-based) of "Total Expenses", so data ends row above
          break;
        }
      }
      
      if (categoryStartRow === -1 || categoryEndRow === -1 || categoryEndRow < categoryStartRow) {
        Logger.log("ChartBuilder: Expense category data range for charts not found in analysis sheet.");
        return; // Cannot create charts if data range is not identified
      }
      
      // Create Pie Chart for Expenditure Breakdown by Category
      // Uses Category (Column A), Type (Column B), Amount (Column C) from the identified range
      const pieDataRange = this.sheet.getRange(categoryStartRow, 1, categoryEndRow - categoryStartRow + 1, 3);
      const pieChart = this.sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(pieDataRange)
        .setPosition(startRow, 1, 0, 0)
        .setOption('title', 'Expenditure Breakdown by Category')
        .setOption('titleTextStyle', { 
          color: this.config.getColors().CHART.TITLE, 
          fontSize: 16, 
          bold: true 
        })
        .setOption('pieSliceText', 'percentage')
        .setOption('pieHole', 0.4)
        .setOption('legend', { 
          position: 'right', 
          textStyle: { 
            color: this.config.getColors().CHART.TEXT, 
            fontSize: 12 
          }
        })
        .setOption('colors', this.config.getColors().CHART.SERIES) // Use predefined chart series colors from config
        .setOption('width', 450)
        .setOption('height', 300)
        .build();
      
      this.sheet.insertChart(pieChart);
      
      // Create Column Chart for Expense Rates vs Targets
      // Uses Category (Col A), % of Income (Col D), Target % (Col E)
      const columnChartDataRanges = [
        this.sheet.getRange(categoryStartRow, 1, categoryEndRow - categoryStartRow + 1, 1), // Categories (X-axis labels)
        this.sheet.getRange(categoryStartRow, 4, categoryEndRow - categoryStartRow + 1, 1), // Actual % of Income (Series 1)
        this.sheet.getRange(categoryStartRow, 5, categoryEndRow - categoryStartRow + 1, 1)  // Target % (Series 2)
      ];
      
      const columnChartBuilder = this.sheet.newChart();
      columnChartDataRanges.forEach(range => columnChartBuilder.addRange(range));
      
      const columnChart = columnChartBuilder
        .setChartType(Charts.ChartType.COLUMN) // Column chart type
        .setPosition(startRow, 5, 0, 0) // Positioned to the right of the pie chart
        .setOption('title', 'Expense Rates vs Targets')
        .setOption('titleTextStyle', { 
          color: this.config.getColors().CHART.TITLE, 
          fontSize: 16, 
          bold: true 
        })
        .setOption('legend', { 
          position: 'top', 
          textStyle: { 
            color: this.config.getColors().CHART.TEXT, 
            fontSize: 12 
          }
        })
        .setOption('colors', [ // Specific colors for actual vs target bars
          this.config.getColors().UI.EXPENSE_FONT || "#FF0000", // Color for actual expense rate bar
          this.config.getColors().UI.INCOME_FONT || "#008000"  // Color for target rate bar (using income font as placeholder)
        ])
        .setOption('width', 450)
        .setOption('height', 300)
        .setOption('hAxis', { 
          title: 'Category', 
          titleTextStyle: { color: this.config.getColors().CHART.TEXT },
          textStyle: { color: this.config.getColors().CHART.TEXT, fontSize: 10 }
        })
        .setOption('vAxis', { 
          title: 'Rate (% of Income)', 
          titleTextStyle: { color: this.config.getColors().CHART.TEXT },
          textStyle: { color: this.config.getColors().CHART.TEXT },
          format: 'percent'
        })
        .setOption('bar', { groupWidth: '75%' })
        .build();
      
      this.sheet.insertChart(columnChart);
    }
  }
  
  /**
   * @classdesc Orchestrates the financial analysis process, using DataExtractor,
   * MetricCardBuilder, and ChartBuilder to generate the analysis sheet.
   * @class FinancialAnalysisService
   * @private
   */
  class FinancialAnalysisService {
    /**
     * Creates an instance of the main FinancialAnalysisService.
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} overviewSheet - The sheet containing overview data.
     */
    constructor(spreadsheet, overviewSheet) {
      /** @type {GoogleAppsScript.Spreadsheet.Spreadsheet} */
      this.spreadsheet = spreadsheet;
      /** @type {GoogleAppsScript.Spreadsheet.Sheet} */
      this.overviewSheet = overviewSheet;
      /** @type {GoogleAppsScript.Spreadsheet.Sheet} */
      this.analysisSheet = utils.getOrCreateSheet(spreadsheet, config.getSection('SHEETS').ANALYSIS);
      /** @type {DataExtractor} */
      this.dataExtractor = new DataExtractor(overviewSheet);
      /** @type {SheetBuilder} */
      this.builder = sheetBuilder.create(this.analysisSheet);
      /** @type {MetricsCalculatorModule} */
      this.metricsCalculator = metricsCalculator;
      /** @type {FormulaBuilderModule} */
      this.formulaBuilder = formulaBuilder;
    }
    
    /**
     * Performs the complete financial analysis, building the analysis sheet.
     * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number}} Result from SheetBuilder's finalize.
     * @memberof FinancialAnalysisService
     */
    analyze() {
      const metrics = this.dataExtractor.extractMetrics();
      
      this.builder
        .clear()
        .addHeaderRow(['Financial Analysis'], { // Main title row
          merge: 6, // Merge across 6 columns (A-F)
          background: config.getSection('COLORS').UI.HEADER_BG,
          fontWeight: 'bold',
          fontColor: config.getSection('COLORS').UI.HEADER_FONT,
          horizontalAlignment: 'center'
        })
        .freezeRows(1); // Freeze the header row
      
      // Define column widths for the analysis sheet layout
      this.builder.setColumnWidths({
        1: 20,   // Column A: Narrow left margin/spacer
        2: 120,  // Column B: First metric card column
        3: 120,  // Column C: First metric card column
        4: 20,   // Column D: Spacer column between card pairs
        5: 120,  // Column E: Second metric card column
        6: 120   // Column F: Second metric card column
      });
      
      this.buildKeyMetricsSection(metrics);
      this.buildExpenseAnalysis(metrics);
      this.createCharts(metrics);
      
      return this.builder.finalize();
    }
    
    /**
     * Builds the key financial metrics section on the analysis sheet.
     * This section displays a series of paired metric cards (e.g., an absolute cash flow value and a related rate/percentage).
     * @param {{
     *   income: ({row: number, total: number, average: number, totalRef: string, averageRef: string, monthlyValuesRangeRef: string}|null),
     *   expenses: ({row: number, total: number, average: number, totalRef: string, averageRef: string, monthlyValuesRangeRef: string}),
     *   savings: ({row: number, total: number, average: number, totalRef: string, averageRef: string, monthlyValuesRangeRef: string}|null),
     *   categories: {income: Array<object>, expenses: Array<object>, savings: Array<object>},
     *   essentials: ({row: number, total: number, average: number, totalRef: string, averageRef: string, monthlyValuesRangeRef: string}|null),
     *   wantsPleasure: ({row: number, total: number, average: number, totalRef: string, averageRef: string, monthlyValuesRangeRef: string}|null),
     *   extra: ({row: number, total: number, average: number, totalRef: string, averageRef: string, monthlyValuesRangeRef: string}|null)
     * }} metrics - An object containing various financial metrics extracted by `DataExtractor`.
     *   Key properties (e.g., `income`, `essentials`) hold objects with `totalRef` and `averageRef`
     *   (cell references from the overview sheet) and other details, or null if not found.
     * @memberof FinancialAnalysisService
     */
    buildKeyMetricsSection(metrics) {
      const totals = metrics;
      const targetRates = config.getSection('TARGET_RATES');

      // Helper function to safely access properties of metric objects (which might be null)
      const getRef = (metric, refType) => metric && metric[refType] ? metric[refType] : null;

      const incomeAvgRef = getRef(totals.income, 'averageRef');
      const incomeTotalRef = getRef(totals.income, 'totalRef');
      const savingsAvgRef = getRef(totals.savings, 'averageRef');
      const savingsTotalRef = getRef(totals.savings, 'totalRef');
      const essentialsAvgRef = getRef(totals.essentials, 'averageRef');
      const essentialsTotalRef = getRef(totals.essentials, 'totalRef');
      const wantsPleasureAvgRef = getRef(totals.wantsPleasure, 'averageRef');
      const wantsPleasureTotalRef = getRef(totals.wantsPleasure, 'totalRef');
      const extraAvgRef = getRef(totals.extra, 'averageRef');
      // const extraTotalRef = getRef(totals.extra, 'totalRef'); // Not used directly in formulas below
      const expensesAvgRef = getRef(totals.expenses, 'averageRef');
      const expensesTotalRef = getRef(totals.expenses, 'totalRef');
      
      // Define configurations for pairs of metric cards.
      // Each object in this array represents a row of two cards on the analysis sheet:
      // - 'cashFlow': Configuration for the left card, typically displaying absolute monetary values.
      // - 'rate': Configuration for the right card, typically displaying a related percentage or rate.
      // Formulas use cell references (e.g., incomeAvgRef) extracted from the overview sheet.
      // IFERROR is used in formulas to gracefully handle cases where references might be invalid or data missing, displaying "N/A" or 0.
      const pairedMetrics = [
        { // Pair 1: Gross Income & Overall Savings Rate
          cashFlow: {
            name: 'Total Gross Income',
            avgFormula: incomeAvgRef ? `=IFERROR(${incomeAvgRef}, "N/A")` : `="N/A"`, 
            totalFormula: incomeTotalRef ? `=IFERROR(${incomeTotalRef}, "N/A")` : `="N/A"`, 
            sparklinePlaceholderText: `[Trend: Gross Income]`,
            valueType: 'currency',
            description: "Total income from all sources before any deductions or allocations."
          },
          rate: { // Corresponding rate card
            name: 'Overall Savings Rate',
            avgFormula: (savingsAvgRef && incomeAvgRef) ? `=IFERROR(${savingsAvgRef}/${incomeAvgRef},0)` : `=0`, 
            targetValue: targetRates.SAVINGS || 0.2, 
            sparklinePlaceholderText: `[Trend: Savings Rate]`,
            valueType: 'percentage',
            avgLabel: 'Avg Rate',
            totalLabel: 'Target',
            description: "Percentage of gross income allocated to savings."
          }
        },
        {
          cashFlow: {
            name: 'Income after Essentials',
            avgFormula: (incomeAvgRef && essentialsAvgRef) ? `=IFERROR(${incomeAvgRef}+${essentialsAvgRef}, "N/A")` : `="N/A"`,
            totalFormula: (incomeTotalRef && essentialsTotalRef) ? `=IFERROR(${incomeTotalRef}+${essentialsTotalRef}, "N/A")` : `="N/A"`,
            sparklinePlaceholderText: `[Trend: NI after Essentials]`,
            valueType: 'currency',
            description: "Income remaining after covering essential living costs."
          },
          rate: {
            name: 'Essentials Spending Rate',
            avgFormula: (essentialsAvgRef && incomeAvgRef) ? `=IFERROR(ABS(${essentialsAvgRef})/${incomeAvgRef},0)` : `=0`,
            targetValue: targetRates.ESSENTIALS || 0.5,
            sparklinePlaceholderText: `[Trend: Essentials Rate]`,
            valueType: 'percentage',
            avgLabel: 'Avg Rate',
            totalLabel: 'Target',
            description: "Percentage of gross income spent on essential needs."
          }
        },
        {
          cashFlow: {
            name: 'Income after Core Spending',
            avgFormula: (incomeAvgRef && essentialsAvgRef && wantsPleasureAvgRef) ? `=IFERROR(${incomeAvgRef}+${essentialsAvgRef}+${wantsPleasureAvgRef}, "N/A")` : `="N/A"`,
            totalFormula: (incomeTotalRef && essentialsTotalRef && wantsPleasureTotalRef) ? `=IFERROR(${incomeTotalRef}+${essentialsTotalRef}+${wantsPleasureTotalRef}, "N/A")` : `="N/A"`,
            sparklinePlaceholderText: `[Trend: NI after Core Spend]`,
            valueType: 'currency',
            description: "Income after essential and regular discretionary (wants/pleasure) costs."
          },
          rate: {
            name: 'Wants/Pleasure Spending Rate',
            avgFormula: (wantsPleasureAvgRef && incomeAvgRef) ? `=IFERROR(ABS(${wantsPleasureAvgRef})/${incomeAvgRef},0)` : `=0`,
            targetValue: targetRates.WANTS || 0.2,
            sparklinePlaceholderText: `[Trend: Wants Rate]`,
            valueType: 'percentage',
            avgLabel: 'Avg Rate',
            totalLabel: 'Target',
            description: "Percentage of gross income spent on wants and pleasure."
          }
        },
        {
          cashFlow: {
            name: 'Allocatable Income',
            // Income + Essentials (negative) + Wants (negative) - Savings (positive, so subtract)
            avgFormula: (incomeAvgRef && essentialsAvgRef && wantsPleasureAvgRef && savingsAvgRef) ? `=IFERROR(${incomeAvgRef}+${essentialsAvgRef}+${wantsPleasureAvgRef}-${savingsAvgRef}, "N/A")` : `="N/A"`,
            totalFormula: (incomeTotalRef && essentialsTotalRef && wantsPleasureTotalRef && savingsTotalRef) ? `=IFERROR(${incomeTotalRef}+${essentialsTotalRef}+${wantsPleasureTotalRef}-${savingsTotalRef}, "N/A")` : `="N/A"`,
            sparklinePlaceholderText: `[Trend: Allocatable Income]`,
            valueType: 'currency',
            description: "Funds for 'Extra' spending or more savings, after core costs & planned savings."
          },
          rate: {
            name: 'Extra Spending Rate',
            avgFormula: (extraAvgRef && incomeAvgRef) ? `=IFERROR(ABS(${extraAvgRef})/${incomeAvgRef},0)` : `=0`,
            targetValue: targetRates.EXTRA || 0.1,
            sparklinePlaceholderText: `[Trend: Extra Rate]`,
            valueType: 'percentage',
            avgLabel: 'Avg Rate',
            totalLabel: 'Target',
            description: "Percentage of gross income spent on non-categorized extra items."
          }
        },
        {
          cashFlow: {
            name: 'Final Net Surplus/Deficit',
            // Income + Expenses (negative) - Savings (positive, so subtract)
            avgFormula: (incomeAvgRef && expensesAvgRef && savingsAvgRef) ? `=IFERROR(${incomeAvgRef}+${expensesAvgRef}-${savingsAvgRef}, "N/A")` : `="N/A"`,
            totalFormula: (incomeTotalRef && expensesTotalRef && savingsTotalRef) ? `=IFERROR(${incomeTotalRef}+${expensesTotalRef}-${savingsTotalRef}, "N/A")` : `="N/A"`,
            sparklinePlaceholderText: `[Trend: Final Surplus/Deficit]`,
            valueType: 'currency',
            description: "The final financial surplus or deficit after all income, expenses, and savings."
          },
          rate: {
            name: 'Net Surplus Rate',
            avgFormula: (incomeAvgRef && expensesAvgRef && savingsAvgRef) ? `=IFERROR((${incomeAvgRef}+${expensesAvgRef}-${savingsAvgRef})/${incomeAvgRef},0)` : `=0`,
            targetValue: 0,
            sparklinePlaceholderText: `[Trend: Surplus Rate]`,
            valueType: 'percentage',
            avgLabel: 'Avg Rate',
            totalLabel: 'Target',
            description: "Final surplus/deficit as a percentage of gross income."
          }
        }
      ];
      
      this.builder
        .addSectionHeader('Key Financial Metrics', {
          merge: 6,
          background: '#D3D3D3',
          fontWeight: 'bold',
          fontSize: 14,
          horizontalAlignment: 'center',
          verticalAlignment: 'middle'
        })
        .setRowHeights({ [this.builder.getCurrentRow() - 1]: 30 }); // Set height for the section header row
      
      // Instantiate MetricCardBuilder and command it to create the grid of cards
      const cardBuilder = new MetricCardBuilder(this.builder, config, this.metricsCalculator);
      cardBuilder.createCardGrid(pairedMetrics, this.builder.getCurrentRow());
    }
    
    /**
     * Builds the expense analysis table on the analysis sheet.
     * @param {ReturnType<DataExtractor['extractMetrics']>} metrics - Extracted metrics from DataExtractor.
     * @memberof FinancialAnalysisService
     */
    buildExpenseAnalysis(metrics) {
      // Filter out subcategories from expenses and sort by average amount (descending) for display
      const categories = metrics.categories.expenses
        .filter(c => !c.subcategory) // We only want main categories for this table
        .sort((a, b) => b.amount - a.amount); // Sort by average amount, highest first
      
      this.builder
        .addSectionHeader('Expense Categories', { // Section title
          merge: 6,
          background: config.getSection('COLORS').UI.HEADER_BG,
          fontWeight: 'bold',
          fontColor: config.getSection('COLORS').UI.HEADER_FONT,
          horizontalAlignment: 'center'
        })
        .addHeaderRow(['Category', 'Type', 'Amount', '% of Income', 'Target %', 'Variance'], { // Table headers
          background: '#F5F5F5', // Light grey for table header
          fontWeight: 'bold',
          horizontalAlignment: 'center'
        });
      
      const categoryData = []; // To hold data rows for batch writing
      const varianceFormulas = []; // To hold formula configurations for batch application
      
      // Process each expense category
      categories.forEach((cat, index) => {
        // Determine the target rate for this category based on its type (Essentials, Wants, etc.)
        let targetRate = config.getSection('TARGET_RATES').DEFAULT; // Default target rate
        if (cat.type === "Essentials") targetRate = config.getSection('TARGET_RATES').ESSENTIALS;
        else if (cat.type === "Wants/Pleasure") targetRate = config.getSection('TARGET_RATES').WANTS;
        else if (cat.type === "Extra") targetRate = config.getSection('TARGET_RATES').EXTRA;
        
        const currentRow = this.builder.getCurrentRow() + index; // Calculate sheet row number for formulas
        
        // Prepare the data array for this category row
        categoryData.push([
          cat.category,
          cat.type,
          cat.amount, // Average monthly amount from overview
          '',         // Placeholder for '% of Income' formula
          targetRate, // Target rate for this category type
          ''          // Placeholder for 'Variance' formula
        ]);
        
        // Store formula configurations to be applied after data is written
        varianceFormulas.push({
          row: currentRow,
          percentFormula: `=IFERROR(C${currentRow}/${metrics.income.averageRef},0)`, // Calculate % of income
          varianceFormula: `=D${currentRow}-E${currentRow}` // Calculate variance: Actual % - Target %
        });
      });
      
      // Add a 'Total Expenses' summary row to the category data
      const totalRow = this.builder.getCurrentRow() + categoryData.length;
      categoryData.push([
        'Total Expenses', // Label
        'All',            // Type
        metrics.expenses.average, // Average total expenses from overview
        '',               // Placeholder for '% of Income' formula
        config.getSection('TARGET_RATES').TOTAL_EXPENSES || 0.8, // Target total expense rate (e.g., 80%)
        ''                // Placeholder for 'Variance' formula
      ]);
      
      // Add formula configuration for the total expenses row
      varianceFormulas.push({
        row: totalRow,
        percentFormula: `=IFERROR(C${totalRow}/${metrics.income.averageRef},0)`,
        varianceFormula: `=D${totalRow}-E${totalRow}`
      });
      
      // Batch write all category data rows to the sheet
      this.builder.addDataRows(categoryData);
      
      // Batch apply the percentage and variance formulas
      varianceFormulas.forEach(vf => {
        this.builder.sheet.getRange(vf.row, 4).setFormula(vf.percentFormula); // Column D for % of Income
        this.builder.sheet.getRange(vf.row, 6).setFormula(vf.varianceFormula); // Column F for Variance
      });
      
      // Get the range of the newly added data for formatting
      const dataRange = this.builder.sheet.getRange(
        this.builder.getCurrentRow() - categoryData.length, // Start row of the category data
        1, // Start column
        categoryData.length, // Number of rows
        6  // Number of columns (A-F)
      );
      
      // Apply currency formatting to the 'Amount' column (Column C)
      utils.formatAsCurrency(
        dataRange.offset(0, 2, categoryData.length, 1), // Offset to Column C
        config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT
      );
      
      // Apply percentage formatting to '% of Income', 'Target %', and 'Variance' columns (D, E, F)
      utils.formatAsPercentage(
        dataRange.offset(0, 3, categoryData.length, 3) // Offset to Column D, 3 columns wide
      );
      
      // Add conditional formatting to highlight positive (over-budget) variance in red
      const rules = [];
      for (let i = 0; i < categoryData.length; i++) {
        const row = this.builder.getCurrentRow() - categoryData.length + i;
        // Rule: If variance (Column F) is greater than 0 (meaning actual % > target %)
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=F${row}>0`) // Condition for positive variance
          .setBackground("#FFCDD2") // Light red background
          .setRanges([this.builder.sheet.getRange(row, 6)]) // Apply to the variance cell
          .build();
        rules.push(rule);
      }
      this.builder.sheet.setConditionalFormatRules(rules); // Apply all conditional formatting rules
      
      // Style the 'Total Expenses' row for emphasis
      const totalRange = this.builder.sheet.getRange(totalRow, 1, 1, 6);
      totalRange
        .setFontWeight('bold')
        .setBackground(config.getSection('COLORS').UI.HEADER_BG) // Use header background color
        .setFontColor(config.getSection('COLORS').UI.HEADER_FONT); // Use header font color
      
      // Add a border around the entire expense analysis table
      dataRange.setBorder(true, true, true, true, true, true); // All borders, internal and external
    }
    
    /**
     * Creates and inserts charts into the analysis sheet.
     * @param {ReturnType<DataExtractor['extractMetrics']>} metrics - Extracted metrics from DataExtractor.
     * @memberof FinancialAnalysisService
     */
    createCharts(metrics) {
      const chartBuilder = new ChartBuilder(this.builder.sheet, config);
      chartBuilder.createExpenseCharts(
        metrics.categories.expenses,
        this.builder.getCurrentRow() + 2
      );
    }
  }
  
  // Public API
  return {
    /**
     * Analyzes financial data from the overview sheet and generates an analysis report.
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} overviewSheet - The sheet containing the financial overview data.
     * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number}|undefined}
     *   The result from the SheetBuilder's finalize method, or undefined if an error occurs.
     * @memberof FinancialPlanner.FinancialAnalysisService
     */
    analyze: function(spreadsheet, overviewSheet) {
      try {
        uiService.showLoadingSpinner("Analyzing financial data...");
        const service = new FinancialAnalysisService(spreadsheet, overviewSheet);
        const result = service.analyze();
        uiService.hideLoadingSpinner();
        return result;
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Error in financial analysis");
        throw error; // Re-throw to allow caller to handle if necessary
      }
    },
    
    /**
     * Public method to trigger the financial analysis and display the key metrics sheet.
     * Handles UI notifications and error reporting.
     * @memberof FinancialPlanner.FinancialAnalysisService
     */
    showKeyMetrics: function() {
      try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const overviewSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').OVERVIEW);
        
        if (!overviewSheet) {
          uiService.showErrorNotification(
            "Error", 
            "Overview sheet not found. Please generate the financial overview first."
          );
          return;
        }
        
        this.analyze(spreadsheet, overviewSheet);
        
        const analysisSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').ANALYSIS);
        if (analysisSheet) {
          analysisSheet.activate();
        }
        
        uiService.showSuccessNotification("Financial analysis complete.");
      } catch (error) {
        errorService.handle(error, "Failed to show key metrics");
      }
    }
  };
})(
  FinancialPlanner.Utils,
  FinancialPlanner.UIService,
  FinancialPlanner.ErrorService,
  FinancialPlanner.Config,
  FinancialPlanner.SheetBuilder,
  FinancialPlanner.MetricsCalculator,
  FinancialPlanner.FormulaBuilder
);

// Backward compatibility
/**
 * Global function to trigger and show the key financial metrics analysis.
 * Delegates to `FinancialPlanner.FinancialAnalysisService.showKeyMetrics()`.
 * @global
 */
function showKeyMetrics() {
  if (typeof FinancialPlanner !== 'undefined' && 
      FinancialPlanner.FinancialAnalysisService && 
      FinancialPlanner.FinancialAnalysisService.showKeyMetrics) {
    FinancialPlanner.FinancialAnalysisService.showKeyMetrics();
  } else {
    Logger.log("Global showKeyMetrics: FinancialPlanner.FinancialAnalysisService not available.");
  }
}
