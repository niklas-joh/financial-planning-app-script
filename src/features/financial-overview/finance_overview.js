/**
 * @fileoverview Financial Overview Generator for the Financial Planning Tools.
 * This module is responsible for creating and managing the main financial overview sheet,
 * which summarizes income, expenses, and savings based on transaction data.
 * It leverages various services for data processing, sheet construction, and UI interactions.
 * Version: 3.0.0
 * @module features/financial-overview
 */

/**
 * @namespace FinancialPlanner.FinanceOverview
 * @description Service responsible for generating and managing the comprehensive financial overview sheet.
 * This IIFE sets up the service with its numerous dependencies.
 * @param {UtilsModule} utils - Instance of the Utils module.
 * @param {UIServiceModule} uiService - Instance of the UI Service module.
 * @param {CacheServiceModule} cacheService - Instance of the Cache Service module.
 * @param {ErrorServiceModule} errorService - Instance of the Error Service module.
 * @param {ConfigModule} config - Instance of the Config module.
 * @param {SettingsServiceModule} settingsService - Instance of the Settings Service module.
 * @param {SheetBuilderModule} sheetBuilder - Instance of the Sheet Builder module (factory).
 * @param {FormulaBuilderModule} formulaBuilder - Instance of the Formula Builder module.
 * @param {DataProcessorModule} dataProcessor - Instance of the Data Processor module.
 * @param {FinancialPlanner.FinancialAnalysisService} analysisService - Instance of the Financial Analysis service.
 */
FinancialPlanner.FinanceOverview = (function(
  utils, uiService, cacheService, errorService, config, settingsService, 
  sheetBuilder, formulaBuilder, dataProcessor, analysisService
) {
  
  /**
   * @classdesc Main builder class for constructing the financial overview sheet.
   * It orchestrates data fetching, processing, and sheet layout using various helper services.
   * @class FinancialOverviewBuilder
   * @private
   */
  class FinancialOverviewBuilder {
    /**
     * Creates an instance of FinancialOverviewBuilder.
     * Initializes properties that will be set during the `initialize` phase.
     */
    constructor() {
      /** @type {GoogleAppsScript.Spreadsheet.Spreadsheet} */
      this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      /** @type {boolean} */
      this.showSubCategories = settingsService.getShowSubCategories();
      /** @type {SheetBuilder|null} */
      this.builder = null;
      /** @type {DataProcessor|null} */
      this.processor = null;
      /** @type {object|null} Processed and grouped transaction data. */
      this.groupedData = null;
      /** @type {object|null} Column indices from the transaction sheet. */
      this.columnIndices = null;
    }
    
    /**
     * Initializes the builder by setting up sheet references, data processor,
     * and fetching/processing necessary data (categories, grouped data).
     * @returns {FinancialOverviewBuilder} The builder instance for chaining.
     * @throws {Error} If the transactions sheet is not found or data structure is invalid.
     * @memberof FinancialOverviewBuilder
     */
    initialize() {
      const sheetNames = config.getSection('SHEETS');
      const overviewSheet = utils.getOrCreateSheet(this.spreadsheet, sheetNames.OVERVIEW);
      this.builder = sheetBuilder.create(overviewSheet);
      
      const transactionSheet = this.spreadsheet.getSheetByName(sheetNames.TRANSACTIONS);
      if (!transactionSheet) {
        throw errorService.create(`Required sheet "${sheetNames.TRANSACTIONS}" not found`, { 
          severity: "high" 
        });
      }
      
      const transactionData = transactionSheet.getDataRange().getValues();
      this.columnIndices = dataProcessor.getColumnIndices(transactionData[0]);
      this.processor = dataProcessor.create(transactionData, this.columnIndices);
      
      // Validate data structure
      this.processor.validateStructure();
      
      // Get cached or process data
      const combinations = cacheService.get('finance_overview_categories', 
        () => this.processor.getUniqueCombinations(this.showSubCategories)
      );
      
      this.groupedData = cacheService.get('finance_overview_grouped',
        () => this.processor.groupByType(combinations)
      );
      
      return this;
    }
    
    /**
     * Builds the entire financial overview sheet structure, including headers,
     * UI controls, data sections (income, expenses, savings), net calculations,
     * and applies final formatting like column widths.
     * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number}}
     *   The result from the SheetBuilder's finalize method, containing the built sheet and last row.
     * @memberof FinancialOverviewBuilder
     */
    build() {
      this.builder
        .clear()
        .addHeaderRow(config.getSection('HEADERS'), {
          background: config.getSection('COLORS').UI.HEADER_BG,
          fontColor: config.getSection('COLORS').UI.HEADER_FONT,
          fontWeight: 'bold',
          horizontalAlignment: 'center',
          verticalAlignment: 'middle'
        })
        .freezeRows(1);
      
      // Add UI controls
      this.setupUIControls();
      
      // Build sections
      this.buildIncomeSection();
      this.buildExpenseSection();
      this.buildSavingsSection();
      this.buildNetCalculations();
      
      // Apply column widths
      const widths = config.getSection('UI').COLUMN_WIDTHS;
      this.builder.setColumnWidths({
        1: widths.TYPE,
        2: widths.CATEGORY,
        3: widths.SUBCATEGORY,
        4: widths.SHARED
      });
      
      // Set month column widths
      for (let i = 5; i <= 16; i++) {
        this.builder.sheet.setColumnWidth(i, widths.MONTH);
      }
      this.builder.sheet.setColumnWidth(17, widths.AVERAGE);
      this.builder.sheet.setColumnWidth(18, widths.AVERAGE);
      
      // Hide subcategory column if needed
      if (!this.showSubCategories) {
        this.builder.sheet.hideColumns(3, 1);
      }
      
      return this.builder.finalize();
    }
    
    /**
     * Sets up UI controls on the overview sheet, such as the "Show Sub-Categories" checkbox.
     * @memberof FinancialOverviewBuilder
     */
    setupUIControls() {
      const uiConfig = config.getSection('UI').SUBCATEGORY_TOGGLE;
      const sheet = this.builder.sheet;
      
      sheet.getRange(uiConfig.LABEL_CELL)
        .setValue(uiConfig.LABEL_TEXT)
        .setFontWeight('bold');
      
      const checkbox = sheet.getRange(uiConfig.CHECKBOX_CELL);
      checkbox.insertCheckboxes()
        .setValue(this.showSubCategories)
        .setNote(uiConfig.NOTE_TEXT);
    }
    
    /**
     * Builds the income section of the overview sheet, including category rows and totals.
     * @memberof FinancialOverviewBuilder
     */
    buildIncomeSection() {
      const transactionTypes = config.getSection('TRANSACTION_TYPES');
      const incomeData = this.groupedData[transactionTypes.INCOME];
      
      if (!incomeData || incomeData.length === 0) return;
      
      this.builder
        .addSectionHeader('Income', {
          merge: 18,
          background: config.getSection('COLORS').UI.SECTION_HEADER_BG || '#d3d3d3',
          fontWeight: 'bold',
          fontSize: 12
        });
      
      const startRow = this.builder.getCurrentRow();
      
      // Add income type row
      this.builder.addDataRows([[transactionTypes.INCOME, '', '', '']], {
        formatting: { fontWeight: 'bold' }
      });
      
      // Add income categories
      const categoryData = incomeData.map(combo => [
        combo.type,
        combo.category,
        combo.subcategory,
        ''
      ]);
      
      const formulas = this.generateRowFormulas(incomeData, startRow + 1);
      
      this.builder.addDataRows(categoryData, {
        formulas: formulas
      });
      
      // Add income total
      this.addTypeTotal('Income', startRow + 1, incomeData.length);
      
      this.builder.addBlankRow();
    }
    
    /**
     * Builds the expense section of the overview sheet.
     * Iterates through configured expense types, adds category rows, and calculates totals.
     * @memberof FinancialOverviewBuilder
     */
    buildExpenseSection() {
      const expenseTypes = config.getSection('EXPENSE_TYPES');
      const colors = config.getSection('COLORS');
      
      const expenseTypeTotalRows = [];
      let hasExpenses = false;
      
      // Check if we have any expenses
      expenseTypes.forEach(type => {
        if (this.groupedData[type] && this.groupedData[type].length > 0) {
          hasExpenses = true;
        }
      });
      
      if (!hasExpenses) return;
      
      this.builder.addSectionHeader('Expenses', {
        merge: 18,
        background: colors.UI.SECTION_HEADER_BG || '#d3d3d3',
        fontWeight: 'bold',
        fontSize: 12
      });
      
      // Add each expense type
      expenseTypes.forEach(type => {
        const typeData = this.groupedData[type];
        if (!typeData || typeData.length === 0) return;
        
        const typeRow = this.builder.getCurrentRow();
        expenseTypeTotalRows.push(typeRow);
        
        // Add type total row with embedded formulas
        this.addTypeRowWithEmbeddedTotals(type, typeRow);
        
        // Add category rows
        const categoryData = typeData.map(combo => [
          combo.type,
          combo.category,
          combo.subcategory,
          ''
        ]);
        
        const startRow = this.builder.getCurrentRow();
        const formulas = this.generateRowFormulas(typeData, startRow);
        
        this.builder.addDataRows(categoryData, {
          formulas: formulas
        });
        
        // Add checkboxes for shared expenses
        for (let i = 0; i < typeData.length; i++) {
          this.builder.sheet.getRange(startRow + i, 4).insertCheckboxes();
        }
        
        this.builder.addBlankRow();
      });
      
      // Add total expenses row
      if (expenseTypeTotalRows.length > 0) {
        this.addTotalExpensesRow(expenseTypeTotalRows);
        this.builder.addBlankRow();
      }
    }
    
    /**
     * Builds the savings section of the overview sheet, including category rows and totals.
     * @memberof FinancialOverviewBuilder
     */
    buildSavingsSection() {
      const transactionTypes = config.getSection('TRANSACTION_TYPES');
      const savingsData = this.groupedData[transactionTypes.SAVINGS];
      
      if (!savingsData || savingsData.length === 0) return;
      
      this.builder.addSectionHeader('Savings', {
        merge: 18,
        background: config.getSection('COLORS').UI.SECTION_HEADER_BG || '#d3d3d3',
        fontWeight: 'bold',
        fontSize: 12
      });
      
      const savingsRow = this.builder.getCurrentRow();
      
      // Add savings type row with embedded totals
      this.addTypeRowWithEmbeddedTotals(transactionTypes.SAVINGS, savingsRow);
      
      // Add savings categories
      const categoryData = savingsData.map(combo => [
        combo.type,
        combo.category,
        combo.subcategory,
        ''
      ]);
      
      const startRow = this.builder.getCurrentRow();
      const formulas = this.generateRowFormulas(savingsData, startRow);
      
      this.builder.addDataRows(categoryData, {
        formulas: formulas
      });
      
      this.builder.addBlankRow();
    }
    
    /**
     * Builds the net calculations section at the bottom of the overview sheet,
     * summarizing key differences (e.g., Income - Expenses).
     * @memberof FinancialOverviewBuilder
     */
    buildNetCalculations() {
      const sheet = this.builder.sheet;
      const data = sheet.getDataRange().getValues();
      const totals = this.findTotalRows(data);
      
      if (!totals.incomeRow || !totals.expensesRow) {
        console.log("Missing required totals for net calculations");
        return;
      }
      
      this.builder.addSectionHeader('Net Calculations', {
        merge: 18,
        background: config.getSection('COLORS').UI.NET_BG,
        fontWeight: 'bold',
        fontColor: config.getSection('COLORS').UI.NET_FONT
      });
      
      // Net (Income - Expenses before Extra)
      if (totals.essentialsRow && totals.wantsPleasureRow) {
        this.addNetCalculationRow(
          'Net (Income - Expenses before Extra)',
          [
            { reference: `${utils.columnToLetter(5)}${totals.incomeRow}`, operation: 'add' },
            { reference: `${utils.columnToLetter(5)}${totals.essentialsRow}`, operation: 'add' },
            { reference: `${utils.columnToLetter(5)}${totals.wantsPleasureRow}`, operation: 'add' }
          ]
        );
      }
      
      // Net (Total Income - Expenses)
      this.addNetCalculationRow(
        'Net (Total Income - Expenses)',
        [
          { reference: `${utils.columnToLetter(5)}${totals.incomeRow}`, operation: 'add' },
          { reference: `${utils.columnToLetter(5)}${totals.expensesRow}`, operation: 'add' }
        ]
      );
      
      // Total Expenses + Savings (if savings exists)
      if (totals.savingsRow) {
        this.addNetCalculationRow(
          'Total Expenses + Savings',
          [
            { reference: `${utils.columnToLetter(5)}${totals.expensesRow}`, operation: 'add' },
            { reference: `${utils.columnToLetter(5)}${totals.savingsRow}`, operation: 'add' }
          ]
        );
        
        // Net (Total Income - Expenses - Savings)
        this.addNetCalculationRow(
          'Net (Total Income - Expenses - Savings)',
          [
            { reference: `${utils.columnToLetter(5)}${totals.incomeRow}`, operation: 'add' },
            { reference: `${utils.columnToLetter(5)}${totals.expensesRow}`, operation: 'add' },
            { reference: `${utils.columnToLetter(5)}${totals.savingsRow}`, operation: 'add' }
          ]
        );
      }
    }
    
    // Helper methods
    /**
     * Generates an array of formula configurations for a set of category/subcategory rows.
     * This includes monthly totals, overall total, and average for each row.
     * @param {Array<{type: string, category: string, subcategory: string}>} combinations - Array of category combination objects.
     * @param {number} startRow - The starting row number (1-based) in the sheet for these combinations.
     * @returns {Array<object>} An array of formula configuration objects suitable for `SheetBuilder.addDataRows`.
     * @memberof FinancialOverviewBuilder
     */
    generateRowFormulas(combinations, startRow) {
      const sheetNames = config.getSection('SHEETS');
      const formulas = [];
      
      // Monthly formulas
      for (let monthCol = 5; monthCol <= 16; monthCol++) {
        const monthFormulas = [];
        for (let i = 0; i < combinations.length; i++) {
          const combo = combinations[i];
          const currentRow = startRow + i;
          const monthDate = new Date(2024, monthCol - 5, 1);
          
          const formula = formulaBuilder.buildCategoryTotalFormula({
            transactionSheet: sheetNames.TRANSACTIONS,
            amountColumn: utils.columnToLetter(this.columnIndices.amount + 1),
            typeColumn: utils.columnToLetter(this.columnIndices.type + 1),
            categoryColumn: utils.columnToLetter(this.columnIndices.category + 1),
            subcategoryColumn: utils.columnToLetter(this.columnIndices.subcategory + 1),
            dateColumn: utils.columnToLetter(this.columnIndices.date + 1),
            typeValue: combo.type,
            categoryValue: combo.category,
            subcategoryValue: combo.subcategory,
            monthDate: monthDate,
            overviewSheetName: sheetNames.OVERVIEW,
            currentRow: currentRow,
            showSubCategories: this.showSubCategories
          });
          
          // Add shared divisor for expense types
          const expenseTypes = config.getSection('EXPENSE_TYPES');
          if (expenseTypes.includes(combo.type)) {
            monthFormulas.push(`(${formula})/IF(D${currentRow}=TRUE, 2, 1)`);
          } else {
            monthFormulas.push(formula);
          }
        }
        formulas.push({
          startColumn: monthCol,
          values: monthFormulas.map(f => [f])
        });
      }
      
      // Total and average formulas
      const totalFormulas = [];
      const averageFormulas = [];
      
      for (let i = 0; i < combinations.length; i++) {
        const row = startRow + i;
        totalFormulas.push([formulaBuilder.buildRowTotalFormula('E', 'P', row)]);
        averageFormulas.push([formulaBuilder.buildRowAverageFormula('E', 'P', row)]);
      }
      
      formulas.push({
        startColumn: 17,
        values: totalFormulas
      });
      
      formulas.push({
        startColumn: 18,
        values: averageFormulas
      });
      
      return formulas;
    }
    
    /**
     * Adds a row for a transaction type (e.g., "Income", "Essentials") that includes
     * embedded formulas for monthly totals, overall total, and average for that type.
     * @param {string} type - The transaction type name.
     * @param {number} row - The row number (1-based) where this type summary should be placed.
     * @memberof FinancialOverviewBuilder
     */
    addTypeRowWithEmbeddedTotals(type, row) {
      const colors = config.getSection('COLORS');
      const sheetNames = config.getSection('SHEETS');
      
      // Add the type name
      this.builder.sheet.getRange(row, 1).setValue(type);
      
      // Style the entire row
      const rowRange = this.builder.sheet.getRange(row, 1, 1, 18);
      rowRange
        .setBackground(colors.UI.TYPE_HEADER_TOTAL_BG || '#f0f0f0')
        .setFontColor(colors.UI.TYPE_HEADER_TOTAL_FONT || '#000000');
      
      // Make the type name bold
      this.builder.sheet.getRange(row, 1).setFontWeight('bold');
      
      // Generate formulas
      const formulas = [];
      for (let monthCol = 5; monthCol <= 16; monthCol++) {
        const monthDate = new Date(2024, monthCol - 5, 1);
        
        const formula = formulaBuilder.buildCategoryTotalFormula({
          transactionSheet: sheetNames.TRANSACTIONS,
          amountColumn: utils.columnToLetter(this.columnIndices.amount + 1),
          typeColumn: utils.columnToLetter(this.columnIndices.type + 1),
          categoryColumn: utils.columnToLetter(this.columnIndices.category + 1),
          subcategoryColumn: utils.columnToLetter(this.columnIndices.subcategory + 1),
          dateColumn: utils.columnToLetter(this.columnIndices.date + 1),
          typeValue: type,
          categoryValue: null,
          subcategoryValue: null,
          monthDate: monthDate,
          overviewSheetName: sheetNames.OVERVIEW,
          currentRow: row,
          showSubCategories: this.showSubCategories
        });
        
        formulas.push(formula);
      }
      
      // Add total and average formulas
      formulas.push(formulaBuilder.buildRowTotalFormula('E', 'P', row));
      formulas.push(formulaBuilder.buildRowAverageFormula('E', 'P', row));
      
      // Apply formulas
      this.builder.sheet.getRange(row, 5, 1, 14).setFormulas([formulas]);
      
      // Format as currency
      const currencyFormat = config.getSection('LOCALE').NUMBER_FORMATS.CURRENCY_DEFAULT;
      this.builder.sheet.getRange(row, 5, 1, 14).setNumberFormat(currencyFormat);
      
      this.builder.setCurrentRow(row + 1);
    }
    
    /**
     * Adds a total row for a specific transaction type (e.g., "Total Income")
     * by summing up its constituent category rows.
     * @param {string} type - The transaction type name (e.g., "Income").
     * @param {number} startRow - The starting row number (1-based) of the categories for this type.
     * @param {number} rowCount - The number of category rows for this type.
     * @memberof FinancialOverviewBuilder
     */
    addTypeTotal(type, startRow, rowCount) {
      const typeColors = this.getTypeColors(type);
      
      this.builder.addSummaryRow(
        `Total ${type}`,
        this.generateTotalFormulas(startRow, rowCount),
        {
          background: typeColors.BG,
          fontColor: typeColors.FONT,
          fontWeight: 'bold'
        }
      );
    }
    
    /**
     * Adds a "Total Expenses" row by summing up the totals of individual expense type rows.
     * @param {Array<number>} expenseTypeRows - An array of row numbers where individual expense type totals are located.
     * @memberof FinancialOverviewBuilder
     */
    addTotalExpensesRow(expenseTypeRows) {
      const formulas = [];
      
      for (let col = 5; col <= 18; col++) {
        const colLetter = utils.columnToLetter(col);
        const components = expenseTypeRows.map(row => ({
          reference: `${colLetter}${row}`,
          operation: 'add'
        }));
        
        formulas.push({
          column: col,
          value: formulaBuilder.buildNetFormula(components)
        });
      }
      
      this.builder.addSummaryRow('Total Expenses', formulas, {
        background: config.getSection('COLORS').UI.GRAND_TOTAL_BG || '#bfbfbf',
        fontColor: config.getSection('COLORS').UI.GRAND_TOTAL_FONT || '#000000',
        fontWeight: 'bold',
        fontSize: 11
      });
    }
    
    /**
     * Adds a row for a net calculation (e.g., "Net Income").
     * Formulas for monthly values, total, and average are generated based on the provided components.
     * @param {string} label - The label for this net calculation row.
     * @param {Array<{reference: string, operation: 'add'|'subtract'}>} components - An array of objects defining
     *   the parts of the net calculation. Each component has a `reference` (e.g., "E5" for the first month's value)
     *   and an `operation`.
     * @memberof FinancialOverviewBuilder
     */
    addNetCalculationRow(label, components) {
      const formulas = [];
      
      for (let col = 5; col <= 18; col++) {
        const colComponents = components.map(comp => ({
          reference: comp.reference.replace(/[A-Z]/, utils.columnToLetter(col)),
          operation: comp.operation
        }));
        
        formulas.push({
          column: col,
          value: formulaBuilder.buildNetFormula(colComponents)
        });
      }
      
      // Create an array of formulas in the format expected by addDataRows
      const formulaValues = [];
      for (let i = 0; i < formulas.length; i++) {
        formulaValues.push(formulas[i].value);
      }
      
      this.builder.addSummaryRow(label, formulas, {
        background: '#F5F5F5',
        fontWeight: 'bold'
      });
    }
    
    /**
     * Generates formula configurations for a total row that sums a range of preceding rows.
     * Used for `addTypeTotal`.
     * @param {number} startRow - The starting row number (1-based) of the range to sum.
     * @param {number} rowCount - The number of rows in the range to sum.
     * @returns {Array<{column: number, value: string}>} An array of formula objects for `SheetBuilder.addSummaryRow`.
     * @memberof FinancialOverviewBuilder
     */
    generateTotalFormulas(startRow, rowCount) {
      const formulas = [];
      const endRow = startRow + rowCount - 1;
      
      for (let col = 5; col <= 18; col++) {
        const colLetter = utils.columnToLetter(col);
        formulas.push({
          column: col,
          value: `=SUM(${colLetter}${startRow}:${colLetter}${endRow})`
        });
      }
      
      return formulas;
    }
    
    /**
     * Finds the row numbers for various key total lines (Income, Expenses, Savings, etc.)
     * by searching the first column of the provided 2D data array.
     * @param {Array<Array<*>>} data - The 2D array of sheet data to search.
     * @returns {{
     *   incomeRow: number|null,
     *   expensesRow: number|null,
     *   savingsRow: number|null,
     *   essentialsRow: number|null,
     *   wantsPleasureRow: number|null,
     *   extraRow: number|null
     * }} An object mapping total labels to their 1-based row numbers, or null if not found.
     * @memberof FinancialOverviewBuilder
     */
    findTotalRows(data) {
      const totals = {
        incomeRow: null,
        expensesRow: null,
        savingsRow: null,
        essentialsRow: null,
        wantsPleasureRow: null,
        extraRow: null
      };
      
      for (let i = 0; i < data.length; i++) {
        const cellValue = data[i][0] ? data[i][0].toString() : "";
        
        if (cellValue === "Total Income") totals.incomeRow = i + 1;
        else if (cellValue === "Total Expenses") totals.expensesRow = i + 1;
        else if (cellValue === "Savings") totals.savingsRow = i + 1;
        else if (cellValue === "Essentials") totals.essentialsRow = i + 1;
        else if (cellValue === "Wants/Pleasure") totals.wantsPleasureRow = i + 1;
        else if (cellValue === "Extra") totals.extraRow = i + 1;
      }
      
      return totals;
    }
    
    /**
     * Retrieves predefined background and font colors for a given transaction type string.
     * Used for styling type headers and total rows.
     * @param {string} type - The transaction type string (e.g., "Income", "Essentials").
     * @returns {{BG: string, FONT: string}} An object with `BG` (background) and `FONT` color hex codes.
     *   Returns default colors if the specific type is not found in the configuration.
     * @memberof FinancialOverviewBuilder
     */
    getTypeColors(type) {
      const typeHeaders = config.getSection('COLORS').TYPE_HEADERS;
      const transactionTypes = config.getSection('TRANSACTION_TYPES');
      
      let colors = typeHeaders.DEFAULT;
      const normalizedType = type.toLowerCase();
      
      for (const key in transactionTypes) {
        if (transactionTypes[key].toLowerCase() === normalizedType && typeHeaders[key]) {
          colors = typeHeaders[key];
          break;
        }
      }
      
      if (normalizedType === "wants/pleasure" && typeHeaders.WANTS_PLEASURE) {
        colors = typeHeaders.WANTS_PLEASURE;
      }
      
      return colors;
    }
  }
  
  // Public API
  return {
    /**
     * Creates or recreates the entire financial overview sheet.
     * This involves initializing data, building all sections, and applying formatting.
     * It also triggers a financial analysis if the service is available.
     * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number}|undefined}
     *   The result from the SheetBuilder's finalize method, or undefined if an error occurs.
     * @memberof FinancialPlanner.FinanceOverview
     * @throws {Error} Propagates errors from underlying services if generation fails.
     */
    create: function() {
      try {
        uiService.showLoadingSpinner("Generating financial overview...");
        cacheService.invalidateAll(); // Ensure fresh data for category combinations
        
        const builder = new FinancialOverviewBuilder();
        const result = builder.initialize().build(); // Initialize and build the sheet
        
        // Optionally, trigger financial analysis if the analysis service is integrated
        if (analysisService && analysisService.analyze) {
          analysisService.analyze(builder.spreadsheet, result.sheet);
        }
        
        uiService.hideLoadingSpinner();
        uiService.showSuccessNotification("Financial overview generated successfully!");
        return result;
      } catch (error) {
        uiService.hideLoadingSpinner();
        errorService.handle(error, "Failed to generate financial overview");
        throw error; // Re-throw to allow caller (e.g., a menu function) to know it failed
      }
    },
    
    /**
     * Handles edit events on the overview sheet, specifically for the subcategory toggle checkbox.
     * If the checkbox state changes, it regenerates the overview.
     * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The edit event object.
     * @memberof FinancialPlanner.FinanceOverview
     */
    handleEdit: function(e) {
      try {
        if (e.range.getSheet().getName() !== config.getSection('SHEETS').OVERVIEW) return;
        
        const subcategoryToggle = config.getSection('UI').SUBCATEGORY_TOGGLE;
        if (e.range.getA1Notation() === subcategoryToggle.CHECKBOX_CELL) {
          const newValue = e.range.getValue();
          settingsService.setShowSubCategories(newValue);
          uiService.showLoadingSpinner("Updating overview based on preference change...");
          this.create();
        }
      } catch (error) {
        errorService.handle(error, "Failed to process change on Overview sheet");
      }
    }
  };
})(
  FinancialPlanner.Utils,
  FinancialPlanner.UIService,
  FinancialPlanner.CacheService,
  FinancialPlanner.ErrorService,
  FinancialPlanner.Config,
  FinancialPlanner.SettingsService,
  FinancialPlanner.SheetBuilder,
  FinancialPlanner.FormulaBuilder,
  FinancialPlanner.DataProcessor,
  FinancialPlanner.FinancialAnalysisService
);

// Backward compatibility
/**
 * Global function to trigger the creation of the financial overview sheet.
 * Delegates to `FinancialPlanner.FinanceOverview.create()`.
 * @global
 * @returns {({sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number}|undefined)} Result of overview creation or undefined on error.
 */
function createFinancialOverview() {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.FinanceOverview && FinancialPlanner.FinanceOverview.create) {
    return FinancialPlanner.FinanceOverview.create();
  }
  Logger.log("Global createFinancialOverview: FinancialPlanner.FinanceOverview or its create method is not available.");
  // Consider a UI alert here if critical and no error handling from caller
}
