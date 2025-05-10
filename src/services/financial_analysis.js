/**
 * Financial Planning Tools - Financial Analysis Service
 * 
 * This module provides analytics functionality for financial data through a dedicated
 * service. It creates a separate Analysis sheet with key metrics, expense category analysis,
 * and visualizations.
 * 
 * Version: 2.2.1
 * Last Updated: 2025-05-10
 */

/**
 * @namespace FinancialPlanner.FinancialAnalysisService
 * @description Service for performing financial analysis based on the data aggregated in the 'Overview' sheet.
 * It generates key metrics, analyzes expense categories against targets, and creates visualizations in a dedicated 'Analysis' sheet.
 */
FinancialPlanner.FinancialAnalysisService = (function(utils, uiService, errorService, config) {
  // ============================================================================
  // PRIVATE IMPLEMENTATION
  // ============================================================================
  
  class FinancialAnalysisService {
    constructor(spreadsheet, overviewSheet, analysisConfig) {
      this.spreadsheet = spreadsheet;
      this.overviewSheet = overviewSheet;
      this.config = analysisConfig;
      this.analysisSheet = utils.getOrCreateSheet(spreadsheet, this.config.SHEETS.ANALYSIS);
      
      let formatString;
      const globalConfigInstance = config; 
      const globalLocale = globalConfigInstance.getLocale ? globalConfigInstance.getLocale() : null;

      if (globalLocale && globalLocale.NUMBER_FORMATS && globalLocale.NUMBER_FORMATS.CURRENCY_DEFAULT) {
          formatString = globalLocale.NUMBER_FORMATS.CURRENCY_DEFAULT;
      } else {
          if (this.config && this.config.LOCALE && this.config.LOCALE.NUMBER_FORMATS && this.config.LOCALE.NUMBER_FORMATS.CURRENCY_DEFAULT) {
              formatString = this.config.LOCALE.NUMBER_FORMATS.CURRENCY_DEFAULT;
          }
      }

      if (!formatString) {
          formatString = '_-[$€-0]* #,##0_-;_-[RED][$€-0]* #,##0_-;* "-";_-@_-'; 
          Logger.log("Warning: CURRENCY_DEFAULT not found in configuration. Using hardcoded default for FinancialAnalysisService.");
      }
      this.currencyFormatDefault = formatString;
      this.data = null;
      this.totals = null;
    }

    initialize() {
      this.extractDataFromOverview();
      this.setupAnalysisSheet();
    }

    analyze() {
      let currentRow = 2; // Start after main sheet header
      currentRow = this.addKeyMetricsSection(currentRow);
      // Spacing is now handled by the return value of addKeyMetricsSection if needed, or can be added here.
      // The +2 was for a general large section spacer. Individual cards have their own spacing.
      // Let's assume addKeyMetricsSection returns the row *after* the last card's spacing.
      currentRow += 1; // Add one more row of general spacing before next section title
      currentRow = this.addExpenseCategoriesSection(currentRow);
      currentRow += 2; // Space after Expense Categories
      this.createExpenditureCharts(currentRow);
    }

    extractDataFromOverview() {
      const overviewData = this.overviewSheet.getDataRange().getValues();
      this.data = { incomeCategories: [], expenseCategories: [], savingsCategories: [], months: [] };
      this.totals = {
        income: { row: -1, total: 0, average: 0, totalRef: '', averageRef: '', monthlyValuesRangeRef: '' },
        expenses: { row: -1, total: 0, average: 0, totalRef: '', averageRef: '' , monthlyValuesRangeRef: ''}, 
        savings: { row: -1, total: 0, average: 0, totalRef: '', averageRef: '', monthlyValuesRangeRef: '' },
        essentials: { row: -1, total: 0, average: 0, totalRef: '', averageRef: '', monthlyValuesRangeRef: '' },
        wantsPleasure: { row: -1, total: 0, average: 0, totalRef: '', averageRef: '', monthlyValuesRangeRef: '' },
        extra: { row: -1, total: 0, average: 0, totalRef: '', averageRef: '', monthlyValuesRangeRef: '' }
      };
      
      const overviewSheetName = `'${this.config.SHEETS.OVERVIEW}'`;
      const monthlyStartColLetter = utils.columnToLetter(5); 
      const monthlyEndColLetter = utils.columnToLetter(16); 
      const totalColLetter = utils.columnToLetter(17); 
      const averageColLetter = utils.columnToLetter(18);

      for (let i = 4; i <= 15; i++) { 
        this.data.months.push(overviewData[0][i]);
      }
      
      for (let i = 0; i < overviewData.length; i++) {
        const rowData = overviewData[i];
        const currentRowNum = i + 1;
        const monthlyRange = `${overviewSheetName}!${monthlyStartColLetter}${currentRowNum}:${monthlyEndColLetter}${currentRowNum}`;
        
        const assignRefs = (category) => {
          category.row = currentRowNum;
          category.total = rowData[16]; 
          category.average = rowData[17];
          category.totalRef = `${overviewSheetName}!${totalColLetter}${currentRowNum}`;
          category.averageRef = `${overviewSheetName}!${averageColLetter}${currentRowNum}`;
          category.monthlyValuesRangeRef = monthlyRange;
        };

        if (rowData[0] === "Total Income") assignRefs(this.totals.income);
        else if (rowData[0] === "Total Essentials") {
          assignRefs(this.totals.essentials);
          if (this.totals.expenses.row === -1) this.totals.expenses.row = currentRowNum;
          this.totals.expenses.total += rowData[16]; this.totals.expenses.average += rowData[17];
        } else if (rowData[0] === "Total Wants/Pleasure") {
          assignRefs(this.totals.wantsPleasure);
          if (this.totals.expenses.row === -1) this.totals.expenses.row = currentRowNum;
          this.totals.expenses.total += rowData[16]; this.totals.expenses.average += rowData[17];
        } else if (rowData[0] === "Total Extra") {
          assignRefs(this.totals.extra);
          if (this.totals.expenses.row === -1) this.totals.expenses.row = currentRowNum;
          this.totals.expenses.total += rowData[16]; this.totals.expenses.average += rowData[17];
        } else if (rowData[0] === "Total Savings") assignRefs(this.totals.savings);
        
        if (rowData[0] === "Income" && rowData[1]) this.data.incomeCategories.push({ category: rowData[1], subcategory: rowData[2] || "", amount: rowData[17], row: i + 1 });
        else if ((rowData[0] === "Essentials" || rowData[0] === "Wants/Pleasure" || rowData[0] === "Extra") && rowData[1]) this.data.expenseCategories.push({ type: rowData[0], category: rowData[1], subcategory: rowData[2] || "", amount: rowData[17], row: i + 1 });
        else if (rowData[0] === "Savings" && rowData[1]) this.data.savingsCategories.push({ category: rowData[1], subcategory: rowData[2] || "", amount: rowData[17], row: i + 1 });
      }
    }

    setupAnalysisSheet() {
      this.analysisSheet.clear(); this.analysisSheet.clearFormats();
      this.analysisSheet.getRange("A1").setValue("Financial Analysis");
      this.analysisSheet.getRange("A1:F1").setBackground(this.config.COLORS.UI.HEADER_BG).setFontWeight("bold").setFontColor(this.config.COLORS.UI.HEADER_FONT).setHorizontalAlignment("center");
      this.analysisSheet.setFrozenRows(1);
      
      this.analysisSheet.setColumnWidth(1, 20); // Narrow column A for spacing or icons later
      this.analysisSheet.setColumnWidth(2, 120); // Card Col B1
      this.analysisSheet.setColumnWidth(3, 120); // Card Col B2
      this.analysisSheet.setColumnWidth(4, 20);  // Narrow column D for spacing
      this.analysisSheet.setColumnWidth(5, 120); // Card Col E1
      this.analysisSheet.setColumnWidth(6, 120); // Card Col E2
      // If more columns are needed for other sections, they can be added/adjusted later.
      this.analysisSheet.setName(this.config.SHEETS.ANALYSIS);
    }

    addKeyMetricsSection(startRow) {
      let currentRowColB = startRow;
      let currentRowColE = startRow;
      const sheet = this.analysisSheet;
      const totals = this.totals;
      const serviceInstance = this; 

      const createMetricCard = (metricConf, cardStartRow, startCardColumn) => {
        const headerRow = cardStartRow;
        const valuesRow = cardStartRow + 1;
        const sparklineRow = cardStartRow + 2;
        const labelsRow = cardStartRow + 3;
        const cardEndRow = labelsRow;
        const cardWidth = 2; // Cards are 2 columns wide

        // Row Heights
        sheet.setRowHeight(headerRow, 25);
        sheet.setRowHeight(valuesRow, 35);
        sheet.setRowHeight(sparklineRow, 30);
        sheet.setRowHeight(labelsRow, 20);

        // 1. Header Row
        sheet.getRange(headerRow, startCardColumn, 1, cardWidth).merge()
          .setValue(metricConf.name)
          .setBackground('#E2EFDA')
          .setFontWeight('bold')
          .setFontColor('black')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');

        // 2. Values Row
        const monthlyValueCell = sheet.getRange(valuesRow, startCardColumn, 1, 1); // Col 1 of card
        const annualValueCell = sheet.getRange(valuesRow, startCardColumn + 1, 1, 1); // Col 2 of card

        monthlyValueCell.setFormula(metricConf.avgFormula)
          .setFontSize(18).setFontColor('#008000')
          .setHorizontalAlignment('center').setVerticalAlignment('middle');

        if (metricConf.totalFormula) {
          annualValueCell.setFormula(metricConf.totalFormula)
            .setFontSize(18).setFontColor('#008000')
            .setHorizontalAlignment('center').setVerticalAlignment('middle');
        } else if (metricConf.targetValue !== undefined) {
          annualValueCell.setValue(metricConf.targetValue)
            .setFontSize(18).setFontColor('#008000')
            .setHorizontalAlignment('center').setVerticalAlignment('middle');
        } else {
          annualValueCell.setValue('') // Blank if no total/target
            .setHorizontalAlignment('center').setVerticalAlignment('middle');
        }
        
        if (metricConf.valueType === 'currency') {
          monthlyValueCell.setNumberFormat(serviceInstance.currencyFormatDefault);
          if (metricConf.totalFormula) annualValueCell.setNumberFormat(serviceInstance.currencyFormatDefault);
        } else if (metricConf.valueType === 'percentage') {
          monthlyValueCell.setNumberFormat('0.00%');
          if (metricConf.totalFormula || metricConf.targetValue !== undefined) {
            if (typeof metricConf.targetValue === 'number') {
                annualValueCell.setNumberFormat('0.00%');
            }
          }
        }

        // 3. Sparkline Row
        sheet.getRange(sparklineRow, startCardColumn, 1, cardWidth).merge()
          .setValue(metricConf.sparklinePlaceholderText || `[Sparkline: ${metricConf.name}]`)
          .setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setFontStyle('italic').setFontColor('#AAAAAA');

        // 4. Labels Row
        sheet.getRange(labelsRow, startCardColumn, 1, 1) // Col 1 of card
          .setValue(metricConf.avgLabel || 'Monthly Avg.')
          .setFontSize(9).setFontColor('#808080')
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
        
        sheet.getRange(labelsRow, startCardColumn + 1, 1, 1) // Col 2 of card
          .setValue(metricConf.totalLabel || 'Annual Total')
          .setFontSize(9).setFontColor('#808080')
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
        
        // Card Border
        sheet.getRange(headerRow, startCardColumn, 4, cardWidth) // 4 rows, cardWidth columns
          .setBorder(true, true, true, true, true, true, '#A9D18E', SpreadsheetApp.BorderStyle.SOLID_THIN);
        
        return cardEndRow + 1; 
      };

      const cashFlowMetricConfigs = [
        {
          name: 'Net Income (after Essentials)',
          avgFormula: `=${totals.income.averageRef}-${totals.essentials.averageRef}`,
          totalFormula: `=${totals.income.totalRef}-${totals.essentials.totalRef}`,
          sparklinePlaceholderText: `[Trend: NI after Essentials]`,
          valueType: 'currency',
        },
        {
          name: 'Discretionary Spending Power',
          avgFormula: `=${totals.income.averageRef}-${totals.essentials.averageRef}-${totals.savings.averageRef}`,
          totalFormula: `=${totals.income.totalRef}-${totals.essentials.totalRef}-${totals.savings.totalRef}`,
          sparklinePlaceholderText: `[Trend: Discretionary Spending]`,
          valueType: 'currency',
        },
        {
          name: 'Overall Net Cash Flow',
          avgFormula: `=${totals.income.averageRef}-(${totals.essentials.averageRef}+${totals.wantsPleasure.averageRef}+${totals.extra.averageRef})`,
          totalFormula: `=${totals.income.totalRef}-(${totals.essentials.totalRef}+${totals.wantsPleasure.totalRef}+${totals.extra.totalRef})`,
          sparklinePlaceholderText: `[Trend: Net Cash Flow]`,
          valueType: 'currency',
        },
        {
          name: 'Net Income after Running Expenses',
          avgFormula: `=(${totals.income.averageRef}-(${totals.essentials.averageRef}+${totals.wantsPleasure.averageRef}))`,
          totalFormula: `=(${totals.income.totalRef}-(${totals.essentials.totalRef}+${totals.wantsPleasure.totalRef}))`,
          sparklinePlaceholderText: `[Trend: NI after Running Exp.]`,
          valueType: 'currency',
        }
      ];

      const rateMetricConfigs = [
        {
          name: 'Savings Rate',
          avgFormula: `=IFERROR(${totals.savings.averageRef}/${totals.income.averageRef},0)`,
          targetValue: serviceInstance.config.TARGET_RATES.SAVINGS !== undefined ? serviceInstance.config.TARGET_RATES.SAVINGS : 'N/A',
          sparklinePlaceholderText: `[Trend: Savings Rate]`,
          valueType: 'percentage',
          avgLabel: 'Avg Rate',
          totalLabel: 'Target'
        },
        {
          name: 'Essentials Rate',
          avgFormula: `=IFERROR(${totals.essentials.averageRef}/${totals.income.averageRef},0)`,
          targetValue: serviceInstance.config.TARGET_RATES.ESSENTIALS !== undefined ? serviceInstance.config.TARGET_RATES.ESSENTIALS : 'N/A',
          sparklinePlaceholderText: `[Trend: Essentials Rate]`,
          valueType: 'percentage',
          avgLabel: 'Avg Rate',
          totalLabel: 'Target'
        },
        {
          name: 'Wants/Pleasure Rate',
          avgFormula: `=IFERROR(${totals.wantsPleasure.averageRef}/${totals.income.averageRef},0)`,
          targetValue: serviceInstance.config.TARGET_RATES.WANTS_PLEASURE !== undefined ? serviceInstance.config.TARGET_RATES.WANTS_PLEASURE : 'N/A',
          sparklinePlaceholderText: `[Trend: Wants Rate]`,
          valueType: 'percentage',
          avgLabel: 'Avg Rate',
          totalLabel: 'Target'
        },
        {
          name: 'Extra Rate',
          avgFormula: `=IFERROR(${totals.extra.averageRef}/${totals.income.averageRef},0)`,
          targetValue: serviceInstance.config.TARGET_RATES.EXTRA !== undefined ? serviceInstance.config.TARGET_RATES.EXTRA : 'N/A',
          sparklinePlaceholderText: `[Trend: Extra Rate]`,
          valueType: 'percentage',
          avgLabel: 'Avg Rate',
          totalLabel: 'Target'
        }
      ];

      // Add Key Metrics Title spanning B-F (or B-C and E-F separately if preferred)
      sheet.getRange(startRow, 2, 1, 5).merge() // Merge B to F for the title "Key Financial Metrics"
           .setValue("Key Financial Metrics")
           .setFontWeight("bold").setFontSize(14)
           .setHorizontalAlignment("center").setVerticalAlignment("middle")
           .setBackground(this.config.COLORS.UI.HEADER_BG || '#D3D3D3') // Use a header-like background
           .setFontColor(this.config.COLORS.UI.HEADER_FONT || '#000000');
      sheet.setRowHeight(startRow, 30); // Height for the title row
      
      let currentTitleRow = startRow + 1; // Actual cards start below this title
      currentRowColB = currentTitleRow;
      currentRowColE = currentTitleRow;


      cashFlowMetricConfigs.forEach(conf => {
        currentRowColB = createMetricCard(conf, currentRowColB, 2); // Start in Col B (2)
        sheet.setRowHeight(currentRowColB, 15); // Spacing row
        currentRowColB++;
      });
      
      rateMetricConfigs.forEach(conf => {
        currentRowColE = createMetricCard(conf, currentRowColE, 5); // Start in Col E (5)
        sheet.setRowHeight(currentRowColE, 15); // Spacing row
        currentRowColE++;
      });
      
      return Math.max(currentRowColB, currentRowColE); // Return the greater of the two current rows
    }

    addExpenseCategoriesSection(startRow) {
      this.analysisSheet.getRange(startRow, 1).setValue("Expense Categories");
      this.analysisSheet.getRange(startRow, 1, 1, 6).setBackground(this.config.COLORS.UI.HEADER_BG).setFontWeight("bold").setFontColor(this.config.COLORS.UI.HEADER_FONT).setHorizontalAlignment("center");
      startRow++;
      this.analysisSheet.getRange(startRow, 1, 1, 6).setValues([["Category", "Type", "Amount", "% of Income", "Target %", "Variance"]]).setBackground("#F5F5F5").setFontWeight("bold").setHorizontalAlignment("center");
      startRow++;
      
      const categoryStartRow = startRow;
      let currentCategoryRow = 0;
      const categoryData = [];
      const conditionalFormatRules = [];
      
      if (this.data.expenseCategories && this.data.expenseCategories.length > 0) {
        const sortedCategories = [...this.data.expenseCategories].filter(c => !c.subcategory).sort((a, b) => b.amount - a.amount);
        
        sortedCategories.forEach(cat => {
          let targetRate = this.config.TARGET_RATES.DEFAULT;
          if (cat.type === "Essentials") targetRate = this.config.TARGET_RATES.ESSENTIALS;
          else if (cat.type === "Wants/Pleasure") targetRate = this.config.TARGET_RATES.WANTS_PLEASURE || this.config.TARGET_RATES.WANTS;
          else if (cat.type === "Extra") targetRate = this.config.TARGET_RATES.EXTRA;
          
          categoryData.push([
            cat.category, cat.type, cat.amount,
            (this.totals.income.averageRef && this.totals.income.average > 0) ? `=C${startRow + currentCategoryRow}/${this.totals.income.averageRef}` : "N/A",
            targetRate,
            `=D${startRow + currentCategoryRow}-E${startRow + currentCategoryRow}`
          ]);
          conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=F${startRow + currentCategoryRow}>0`).setBackground("#FFCDD2").setRanges([this.analysisSheet.getRange(startRow + currentCategoryRow, 6)]).build());
          currentCategoryRow++;
        });
        
        categoryData.push([
          "Total Expenses", "All", this.totals.expenses.average,
          (this.totals.income.averageRef && this.totals.income.average > 0) ? `=C${startRow + currentCategoryRow}/${this.totals.income.averageRef}` : "N/A",
          0.8, // Example target for total expenses
          `=D${startRow + currentCategoryRow}-E${startRow + currentCategoryRow}`
        ]);
        currentCategoryRow++;
        
        if (categoryData.length > 0) {
          const dataRange = this.analysisSheet.getRange(startRow, 1, categoryData.length, 6);
          dataRange.setValues(categoryData).setBackground(null).setBorder(true, true, true, true, true, true, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID_THIN);
          utils.formatAsCurrency(this.analysisSheet.getRange(startRow, 3, categoryData.length, 1), this.currencyFormatDefault);
          utils.formatAsPercentage(this.analysisSheet.getRange(startRow, 4, categoryData.length, 3)); 
          
          const rules = this.analysisSheet.getConditionalFormatRules();
          conditionalFormatRules.forEach(rule => rules.push(rule));
          this.analysisSheet.setConditionalFormatRules(rules);
          
          this.analysisSheet.getRange(startRow + categoryData.length - 1, 1, 1, 6).setFontWeight("bold").setFontColor(this.config.COLORS.UI.HEADER_FONT).setBackground(this.config.COLORS.UI.HEADER_BG);
        }
      }
      return startRow + currentCategoryRow;
    }

    createExpenditureCharts(startRow) {
      if (!this.data || !this.data.expenseCategories || this.data.expenseCategories.length === 0) return; 
      const analysisData = this.analysisSheet.getDataRange().getValues();
      let categoryStartRow = -1, categoryEndRow = -1;
      for (let i = 0; i < analysisData.length; i++) {
        if (analysisData[i][0] === "Category" && analysisData[i][1] === "Type") categoryStartRow = i + 2; // Data starts on the row after header
        else if (analysisData[i][0] === "Total Expenses" && analysisData[i][1] === "All") { categoryEndRow = i; break; } // Data ends on the row before total
      }

      if (categoryStartRow === -1 || categoryEndRow === -1 || categoryEndRow < categoryStartRow) {
        Logger.log("Expenditure chart: Category data range not found or invalid.");
        return;
      }
      
      // Pie chart for categories (excluding total)
      const pieDataRange = this.analysisSheet.getRange(categoryStartRow, 1, categoryEndRow - categoryStartRow + 1, 3); // Category name and Amount
      const pieChart = this.analysisSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(pieDataRange)
        .setPosition(startRow, 1, 0, 0).setOption('title', 'Expenditure Breakdown by Category')
        .setOption('titleTextStyle', { color: this.config.COLORS.CHART.TITLE, fontSize: 16, bold: true })
        .setOption('pieSliceText', 'percentage').setOption('pieHole', 0.4)
        .setOption('legend', { position: 'right', textStyle: { color: this.config.COLORS.CHART.TEXT, fontSize: 12 }})
        .setOption('colors', this.config.COLORS.CHART.SERIES).setOption('width', 450).setOption('height', 300)
        .setOption('is3D', false).setOption('pieSliceTextStyle', { color: '#FFFFFF', fontSize: 14, bold: true })
        .setOption('tooltip', { showColorCode: true, textStyle: { fontSize: 12 }}).build();
      this.analysisSheet.insertChart(pieChart);

      // Column chart for rates vs targets (excluding total)
      // Range: Category (col A), % of Income (col D), Target % (col E)
      const columnChartDataRanges = [
        this.analysisSheet.getRange(categoryStartRow, 1, categoryEndRow - categoryStartRow + 1, 1), // Category names
        this.analysisSheet.getRange(categoryStartRow, 4, categoryEndRow - categoryStartRow + 1, 1), // % of Income
        this.analysisSheet.getRange(categoryStartRow, 5, categoryEndRow - categoryStartRow + 1, 1)  // Target %
      ];
      
      const columnChartBuilder = this.analysisSheet.newChart().setChartType(Charts.ChartType.COLUMN);
      columnChartDataRanges.forEach(range => columnChartBuilder.addRange(range));
      
      const columnChart = columnChartBuilder.setPosition(startRow, 5, 0, 0) // Positioned to the right of pie, adjust col index if needed
        .setOption('title', 'Expense Rates vs Targets')
        .setOption('titleTextStyle', { color: this.config.COLORS.CHART.TITLE, fontSize: 16, bold: true })
        .setOption('legend', { position: 'top', textStyle: { color: this.config.COLORS.CHART.TEXT, fontSize: 12 }})
        .setOption('colors', [this.config.COLORS.UI.EXPENSE_FONT || "#FF0000", this.config.COLORS.UI.INCOME_FONT || "#008000"])
        .setOption('width', 450).setOption('height', 300)
        .setOption('hAxis', { title: 'Category', titleTextStyle: {color: this.config.COLORS.CHART.TEXT}, textStyle: {color: this.config.COLORS.CHART.TEXT, fontSize: 10}})
        .setOption('vAxis', { title: 'Rate (% of Income)', titleTextStyle: {color: this.config.COLORS.CHART.TEXT}, textStyle: {color: this.config.COLORS.CHART.TEXT}, format: 'percent' })
        .setOption('bar', {groupWidth: '75%'}).setOption('isStacked', false).build();
      this.analysisSheet.insertChart(columnChart);
    }

    suggestSavingsOpportunities() { uiService.showInfoNotification("Info", "suggestSavingsOpportunities called."); }
    detectSpendingAnomalies() { uiService.showInfoNotification("Info", "detectSpendingAnomalies called."); }
    analyzeFixedVsVariableExpenses() { uiService.showInfoNotification("Info", "analyzeFixedVsVariableExpenses called."); }
    generateCashFlowForecast() { uiService.showInfoNotification("Info", "generateCashFlowForecast called."); }
  }
  
  // PUBLIC API
  return {
    analyze: function(spreadsheet, overviewSheet) {
      try {
        uiService.showLoadingSpinner("Analyzing financial data...");
        const analysisConfig = { ...config.get(), TARGET_RATES: { ...config.getSection('TARGET_RATES'), WANTS_PLEASURE: config.getSection('TARGET_RATES').WANTS } };
        const service = new FinancialAnalysisService(spreadsheet, overviewSheet, analysisConfig);
        service.initialize(); 
        service.analyze();    
        uiService.hideLoadingSpinner();
        return service; 
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Error in financial analysis service");
        throw error;
      }
    },
    showKeyMetrics: function() {
      try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const overviewSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').OVERVIEW);
        if (!overviewSheet) {
          uiService.showErrorNotification("Error", "Overview sheet not found. Please generate the financial overview first.");
          return;
        }
        FinancialPlanner.FinancialAnalysisService.analyze(spreadsheet, overviewSheet); 
        const analysisSheet = spreadsheet.getSheetByName(config.getSection('SHEETS').ANALYSIS);
        if (analysisSheet) { 
            analysisSheet.activate();
        }
        uiService.showSuccessNotification("Financial analysis complete.");
      } catch (error) {
        errorService.handle(error, "Failed to show key metrics");
      }
    },
    suggestSavingsOpportunities: function() { uiService.showInfoNotification("Info", "Suggest Savings Opportunities feature called."); },
    detectSpendingAnomalies: function() { uiService.showInfoNotification("Info", "Detect Spending Anomalies feature called."); }
  };
})(FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.ErrorService, FinancialPlanner.Config);

function showKeyMetrics() {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.FinancialAnalysisService && FinancialPlanner.FinancialAnalysisService.showKeyMetrics) {
    FinancialPlanner.FinancialAnalysisService.showKeyMetrics();
  } else {
     Logger.log("Global showKeyMetrics: FinancialPlanner.FinancialAnalysisService not available.");
  }
}
