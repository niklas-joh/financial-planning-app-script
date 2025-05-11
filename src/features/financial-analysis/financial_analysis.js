/**
 * Financial Planning Tools - Financial Analysis Service
 * Version: 3.0.0
 * 
 * This module provides analytics functionality for financial data.
 * Refactored to use new services for better maintainability.
 */

FinancialPlanner.FinancialAnalysisService = (function(
  utils, uiService, errorService, config, 
  sheetBuilder, metricsCalculator, formulaBuilder
) {
  
  /**
   * Data extractor for overview sheet
   */
  class DataExtractor {
    constructor(overviewSheet) {
      this.sheet = overviewSheet;
      this.data = this.sheet.getDataRange().getValues();
      this.overviewSheetName = `'${config.getSection('SHEETS').OVERVIEW}'`;
    }
    
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
    
    findExpenseRows() {
      let totalExpenses = { row: -1, total: 0, average: 0 };
      
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
    
    extractCategories() {
      const categories = {
        income: [],
        expenses: [],
        savings: []
      };
      
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
        } else if (config.getSection('EXPENSE_TYPES').includes(type)) {
          categories.expenses.push(categoryData);
        }
      }
      
      return categories;
    }
  }
  
  /**
   * Metric card builder
   */
  class MetricCardBuilder {
    constructor(builder, config, calculator) {
      this.builder = builder;
      this.config = config;
      this.calculator = calculator;
    }
    
    createCardGrid(metrics, startRow) {
      this.builder.setCurrentRow(startRow);
      
      // Create two columns of cards
      metrics.forEach((pair, index) => {
        this.createCardPair(pair.cashFlow, pair.rate);
        // Add spacing between card pairs
        this.builder.addBlankRow(15);
      });
    }
    
    createCardPair(cashFlowMetric, rateMetric) {
      const currentRow = this.builder.getCurrentRow();
      
      // Create cash flow card (columns B-C)
      this.createCard(cashFlowMetric, currentRow, 2);
      
      // Create rate card (columns E-F)
      this.createCard(rateMetric, currentRow, 5);
      
      this.builder.setCurrentRow(currentRow + 5);
    }
    
    createCard(metric, startRow, startColumn) {
      const sheet = this.builder.sheet;
      
      // Set row heights for card structure
      sheet.setRowHeight(startRow, 25);      // Header
      sheet.setRowHeight(startRow + 1, 35);  // Values
      sheet.setRowHeight(startRow + 2, 30);  // Sparkline
      sheet.setRowHeight(startRow + 3, 20);  // Labels
      sheet.setRowHeight(startRow + 4, 35);  // Description
      
      // Header
      sheet.getRange(startRow, startColumn, 1, 2).merge()
        .setValue(metric.name)
        .setBackground('#E2EFDA')
        .setFontWeight('bold')
        .setFontColor('black')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Monthly value
      const monthlyValueCell = sheet.getRange(startRow + 1, startColumn);
      monthlyValueCell.setFormula(metric.avgFormula)
        .setFontSize(18)
        .setFontColor('#008000')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Annual/Target value
      const annualValueCell = sheet.getRange(startRow + 1, startColumn + 1);
      if (metric.totalFormula) {
        annualValueCell.setFormula(metric.totalFormula);
      } else if (metric.targetValue !== undefined) {
        annualValueCell.setValue(metric.targetValue);
      }
      annualValueCell.setFontSize(18)
        .setFontColor('#008000')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Apply number formatting
      if (metric.valueType === 'currency') {
        const currencyFormat = this.config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT;
        monthlyValueCell.setNumberFormat(currencyFormat);
        if (metric.totalFormula) annualValueCell.setNumberFormat(currencyFormat);
      } else if (metric.valueType === 'percentage') {
        monthlyValueCell.setNumberFormat('0.00%');
        if (typeof metric.targetValue === 'number') {
          annualValueCell.setNumberFormat('0.00%');
        }
      }
      
      // Sparkline placeholder
      sheet.getRange(startRow + 2, startColumn, 1, 2).merge()
        .setValue(metric.sparklinePlaceholderText || `[Sparkline: ${metric.name}]`)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setFontStyle('italic')
        .setFontColor('#AAAAAA');
      
      // Labels
      sheet.getRange(startRow + 3, startColumn)
        .setValue(metric.avgLabel || 'Monthly Avg.')
        .setFontSize(9)
        .setFontColor('#808080')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      sheet.getRange(startRow + 3, startColumn + 1)
        .setValue(metric.totalLabel || (metric.targetValue !== undefined ? 'Target' : 'Annual Total'))
        .setFontSize(9)
        .setFontColor('#808080')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Description
      sheet.getRange(startRow + 4, startColumn, 1, 2).merge()
        .setValue(metric.description || '')
        .setFontSize(9)
        .setFontColor('#595959')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('top')
        .setWrap(true);
      
      // Border for the entire card
      sheet.getRange(startRow, startColumn, 5, 2)
        .setBorder(true, true, true, true, true, true, '#A9D18E', SpreadsheetApp.BorderStyle.SOLID_THIN);
    }
  }
  
  /**
   * Chart builder
   */
  class ChartBuilder {
    constructor(sheet, config) {
      this.sheet = sheet;
      this.config = config;
    }
    
    createExpenseCharts(categoryData, startRow) {
      // Find the expense category data range
      const analysisData = this.sheet.getDataRange().getValues();
      let categoryStartRow = -1;
      let categoryEndRow = -1;
      
      for (let i = 0; i < analysisData.length; i++) {
        if (analysisData[i][0] === "Category" && analysisData[i][1] === "Type") {
          categoryStartRow = i + 2;
        } else if (analysisData[i][0] === "Total Expenses" && analysisData[i][1] === "All") {
          categoryEndRow = i;
          break;
        }
      }
      
      if (categoryStartRow === -1 || categoryEndRow === -1 || categoryEndRow < categoryStartRow) {
        Logger.log("Chart data range not found");
        return;
      }
      
      // Create pie chart
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
        .setOption('colors', this.config.getColors().CHART.SERIES)
        .setOption('width', 450)
        .setOption('height', 300)
        .build();
      
      this.sheet.insertChart(pieChart);
      
      // Create column chart
      const columnChartDataRanges = [
        this.sheet.getRange(categoryStartRow, 1, categoryEndRow - categoryStartRow + 1, 1), // Categories
        this.sheet.getRange(categoryStartRow, 4, categoryEndRow - categoryStartRow + 1, 1), // % of Income
        this.sheet.getRange(categoryStartRow, 5, categoryEndRow - categoryStartRow + 1, 1)  // Target %
      ];
      
      const columnChartBuilder = this.sheet.newChart();
      columnChartDataRanges.forEach(range => columnChartBuilder.addRange(range));
      
      const columnChart = columnChartBuilder
        .setChartType(Charts.ChartType.COLUMN)
        .setPosition(startRow, 5, 0, 0)
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
        .setOption('colors', [
          this.config.getColors().UI.EXPENSE_FONT || "#FF0000",
          this.config.getColors().UI.INCOME_FONT || "#008000"
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
   * Main analysis service
   */
  class FinancialAnalysisService {
    constructor(spreadsheet, overviewSheet) {
      this.spreadsheet = spreadsheet;
      this.overviewSheet = overviewSheet;
      this.analysisSheet = utils.getOrCreateSheet(spreadsheet, config.getSection('SHEETS').ANALYSIS);
      this.dataExtractor = new DataExtractor(overviewSheet);
      this.builder = sheetBuilder.create(this.analysisSheet);
      this.metricsCalculator = metricsCalculator;
      this.formulaBuilder = formulaBuilder;
    }
    
    analyze() {
      const metrics = this.dataExtractor.extractMetrics();
      
      this.builder
        .clear()
        .addHeaderRow(['Financial Analysis'], {
          merge: 6,
          background: config.getSection('COLORS').UI.HEADER_BG,
          fontWeight: 'bold',
          fontColor: config.getSection('COLORS').UI.HEADER_FONT,
          horizontalAlignment: 'center'
        })
        .freezeRows(1);
      
      // Set column widths
      this.builder.setColumnWidths({
        1: 20,   // Narrow left margin
        2: 120,  // Card column
        3: 120,  // Card column
        4: 20,   // Spacer
        5: 120,  // Card column
        6: 120   // Card column
      });
      
      this.buildKeyMetricsSection(metrics);
      this.buildExpenseAnalysis(metrics);
      this.createCharts(metrics);
      
      return this.builder.finalize();
    }
    
    buildKeyMetricsSection(metrics) {
      const totals = metrics;
      const targetRates = config.getSection('TARGET_RATES');

      // Helper to safely get refs
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
      
      const pairedMetrics = [
        {
          cashFlow: {
            name: 'Total Gross Income',
            avgFormula: incomeAvgRef ? `=IFERROR(${incomeAvgRef}, "N/A")` : `="N/A"`,
            totalFormula: incomeTotalRef ? `=IFERROR(${incomeTotalRef}, "N/A")` : `="N/A"`,
            sparklinePlaceholderText: `[Trend: Gross Income]`,
            valueType: 'currency',
            description: "Total income from all sources before any deductions or allocations."
          },
          rate: {
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
        .setRowHeights({ [this.builder.getCurrentRow() - 1]: 30 });
      
      const cardBuilder = new MetricCardBuilder(this.builder, config, this.metricsCalculator);
      cardBuilder.createCardGrid(pairedMetrics, this.builder.getCurrentRow());
    }
    
    buildExpenseAnalysis(metrics) {
      const categories = metrics.categories.expenses
        .filter(c => !c.subcategory)
        .sort((a, b) => b.amount - a.amount);
      
      this.builder
        .addSectionHeader('Expense Categories', {
          merge: 6,
          background: config.getSection('COLORS').UI.HEADER_BG,
          fontWeight: 'bold',
          fontColor: config.getSection('COLORS').UI.HEADER_FONT,
          horizontalAlignment: 'center'
        })
        .addHeaderRow(['Category', 'Type', 'Amount', '% of Income', 'Target %', 'Variance'], {
          background: '#F5F5F5',
          fontWeight: 'bold',
          horizontalAlignment: 'center'
        });
      
      const categoryData = [];
      const varianceFormulas = [];
      
      categories.forEach((cat, index) => {
        let targetRate = config.getSection('TARGET_RATES').DEFAULT;
        if (cat.type === "Essentials") targetRate = config.getSection('TARGET_RATES').ESSENTIALS;
        else if (cat.type === "Wants/Pleasure") targetRate = config.getSection('TARGET_RATES').WANTS;
        else if (cat.type === "Extra") targetRate = config.getSection('TARGET_RATES').EXTRA;
        
        const currentRow = this.builder.getCurrentRow() + index;
        
        categoryData.push([
          cat.category,
          cat.type,
          cat.amount,
          '', // Will be replaced by formula
          targetRate,
          ''  // Will be replaced by formula
        ]);
        
        varianceFormulas.push({
          row: currentRow,
          percentFormula: `=IFERROR(C${currentRow}/${metrics.income.averageRef},0)`,
          varianceFormula: `=D${currentRow}-E${currentRow}`
        });
      });
      
      // Add total row
      const totalRow = this.builder.getCurrentRow() + categoryData.length;
      categoryData.push([
        'Total Expenses',
        'All',
        metrics.expenses.average,
        '', // Will be replaced by formula
        0.8,
        ''  // Will be replaced by formula
      ]);
      
      varianceFormulas.push({
        row: totalRow,
        percentFormula: `=IFERROR(C${totalRow}/${metrics.income.averageRef},0)`,
        varianceFormula: `=D${totalRow}-E${totalRow}`
      });
      
      // Add the data
      this.builder.addDataRows(categoryData);
      
      // Apply formulas
      varianceFormulas.forEach(vf => {
        this.builder.sheet.getRange(vf.row, 4).setFormula(vf.percentFormula);
        this.builder.sheet.getRange(vf.row, 6).setFormula(vf.varianceFormula);
      });
      
      // Format the ranges
      const dataRange = this.builder.sheet.getRange(
        this.builder.getCurrentRow() - categoryData.length,
        1,
        categoryData.length,
        6
      );
      
      // Apply currency format to amount column
      utils.formatAsCurrency(
        dataRange.offset(0, 2, categoryData.length, 1),
        config.getLocale().NUMBER_FORMATS.CURRENCY_DEFAULT
      );
      
      // Apply percentage format to columns D, E, F
      utils.formatAsPercentage(
        dataRange.offset(0, 3, categoryData.length, 3)
      );
      
      // Add conditional formatting for variance
      const rules = [];
      for (let i = 0; i < categoryData.length; i++) {
        const row = this.builder.getCurrentRow() - categoryData.length + i;
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=F${row}>0`)
          .setBackground("#FFCDD2")
          .setRanges([this.builder.sheet.getRange(row, 6)])
          .build();
        rules.push(rule);
      }
      this.builder.sheet.setConditionalFormatRules(rules);
      
      // Style the total row
      const totalRange = this.builder.sheet.getRange(totalRow, 1, 1, 6);
      totalRange
        .setFontWeight('bold')
        .setBackground(config.getSection('COLORS').UI.HEADER_BG)
        .setFontColor(config.getSection('COLORS').UI.HEADER_FONT);
      
      // Add border
      dataRange.setBorder(true, true, true, true, true, true);
    }
    
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
        throw error;
      }
    },
    
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
function showKeyMetrics() {
  if (typeof FinancialPlanner !== 'undefined' && 
      FinancialPlanner.FinancialAnalysisService && 
      FinancialPlanner.FinancialAnalysisService.showKeyMetrics) {
    FinancialPlanner.FinancialAnalysisService.showKeyMetrics();
  } else {
    Logger.log("Global showKeyMetrics: FinancialPlanner.FinancialAnalysisService not available.");
  }
}
