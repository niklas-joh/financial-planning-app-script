/**
 * Financial Planning Tools - Financial Overview Generator
 * 
 * This module creates a comprehensive financial overview sheet based on transaction data.
 * It generates a complete overview with dynamic categories and optional sub-category display
 * based on user preference.
 * 
 * Version: 2.1.0
 * Last Updated: 2025-05-08
 */

/**
 * @namespace FinancialPlanner.FinanceOverview
 * @description Service responsible for generating a comprehensive financial overview sheet.
 */
FinancialPlanner.FinanceOverview = (function(utils, uiService, cacheService, errorService, config, settingsService, analysisServiceInstance) {
  // ============================================================================
  // PRIVATE IMPLEMENTATION
  // ============================================================================
  
  function getProcessedTransactionData(sheet) {
    const rawData = sheet.getDataRange().getValues();
    const headers = rawData[0];
    const indices = {
      type: headers.indexOf("Type"), category: headers.indexOf("Category"),
      subcategory: headers.indexOf("Sub-Category"), date: headers.indexOf("Date"),
      amount: headers.indexOf("Amount"), shared: headers.indexOf("Shared")
    };
    const requiredColumns = ["type", "category", "subcategory", "date", "amount"];
    const missingColumns = requiredColumns.filter(col => indices[col] < 0);
    if (missingColumns.length > 0) {
      throw errorService.create(`Required columns not found: ${missingColumns.join(", ")}`, { severity: "high", headers });
    }
    return { data: rawData, indices };
  }
  
  function getUniqueCategoryCombinations(data, typeCol, categoryCol, subcategoryCol, showSubCategories) {
    const seen = new Set();
    return data.slice(1).reduce((combinations, row) => {
      const type = row[typeCol]; const category = row[categoryCol];
      const subcategory = showSubCategories ? row[subcategoryCol] : "";
      if (!type || !category) return combinations;
      const key = `${type}|${category}|${subcategory || ""}`;
      if (!seen.has(key)) {
        seen.add(key);
        combinations.push({ type, category, subcategory: subcategory || "" });
      }
      return combinations;
    }, []);
  }
  
  function groupCategoryCombinations(combinations) {
    const grouped = combinations.reduce((acc, combo) => {
      if (!acc[combo.type]) acc[combo.type] = [];
      acc[combo.type].push(combo);
      return acc;
    }, {});
    Object.keys(grouped).forEach(type => {
      grouped[type].sort((a, b) => {
        const catComp = a.category.localeCompare(b.category);
        return catComp !== 0 ? catComp : (a.subcategory || "").localeCompare(b.subcategory || "");
      });
    });
    return grouped;
  }
  
  function buildMonthlySumFormula(params, overviewSheetCurrentRow) {
    const { type, category, subcategory, monthDate, sheetName, typeCol, categoryCol, subcategoryCol, dateCol, amountCol } = params;
    const month = monthDate.getMonth() + 1;
    const year = monthDate.getFullYear();
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);
    const startDateFormatted = formatDate(startDate);
    const endDateFormatted = formatDate(endDate);
    const sumRange = `${sheetName}!${utils.columnToLetter(amountCol)}:${utils.columnToLetter(amountCol)}`;
    
    const baseCriteria = [
      `${sheetName}!${utils.columnToLetter(typeCol)}:${utils.columnToLetter(typeCol)}, "${type}"`,
      `${sheetName}!${utils.columnToLetter(dateCol)}:${utils.columnToLetter(dateCol)}, ">=${startDateFormatted}"`,
      `${sheetName}!${utils.columnToLetter(dateCol)}:${utils.columnToLetter(dateCol)}, "<=${endDateFormatted}"`
    ];

    if (category) {
      baseCriteria.push(`${sheetName}!${utils.columnToLetter(categoryCol)}:${utils.columnToLetter(categoryCol)}, "${category}"`);
      if (subcategory) {
        baseCriteria.push(`${sheetName}!${utils.columnToLetter(subcategoryCol)}:${utils.columnToLetter(subcategoryCol)}, "${subcategory}"`);
      }
    }
    
    const sumifsFormula = `SUMIFS(${sumRange}, ${baseCriteria.join(", ")})`;
    const expenseTypes = config.getSection('EXPENSE_TYPES');
    if (category && expenseTypes.includes(type)) {
      const divisorFormula = `IF(D${overviewSheetCurrentRow}=TRUE, 2, 1)`;
      return `(${sumifsFormula}) / ${divisorFormula}`;
    }
    return sumifsFormula;
  }
  
  function formatDate(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), config.getSection('LOCALE').DATE_FORMAT);
  }
  
  function clearSheetContent(sheet) {
    sheet.clear(); sheet.clearFormats(); sheet.getRange("A1:Z1000").setDataValidation(null);
  }
  
  function setupHeaderRow(sheet, showSubCategories) {
    const headers = config.getSection('HEADERS');
    const uiConfig = config.getSection('UI');
    const colors = config.getSection('COLORS').UI;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(colors.HEADER_BG).setFontWeight("bold").setFontColor(colors.HEADER_FONT)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    const { SUBCATEGORY_TOGGLE } = uiConfig;
    sheet.getRange(SUBCATEGORY_TOGGLE.LABEL_CELL).setValue(SUBCATEGORY_TOGGLE.LABEL_TEXT).setFontWeight("bold");
    const checkbox = sheet.getRange(SUBCATEGORY_TOGGLE.CHECKBOX_CELL);
    checkbox.insertCheckboxes().setValue(showSubCategories).setNote(SUBCATEGORY_TOGGLE.NOTE_TEXT);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(4, uiConfig.COLUMN_WIDTHS.SHARED);
  }

  function addMajorSectionHeader(sheet, title, rowIndex) {
    const headersConfig = config.getSection('HEADERS');
    const uiColors = config.getSection('COLORS').UI;
    const style = {
      fontSize: 12, fontWeight: "bold",
      backgroundColor: uiColors.SECTION_HEADER_BG || "#d3d3d3",
      fontColor: uiColors.SECTION_HEADER_FONT || "#000000"
    };
    sheet.getRange(rowIndex, 1).setValue(title);
    const headerRange = sheet.getRange(rowIndex, 1, 1, headersConfig.length);
    headerRange.setBackground(style.backgroundColor).setFontWeight(style.fontWeight)
      .setFontColor(style.fontColor).setFontSize(style.fontSize).setVerticalAlignment("middle");
    return rowIndex + 1;
  }
  
  function _addSimpleTypeLabelRow(sheet, type, rowIndex) {
    const headers = config.getSection('HEADERS');
    sheet.getRange(rowIndex, 1).setValue(type); 
    const rowRange = sheet.getRange(rowIndex, 1, 1, headers.length);
    rowRange.setFontWeight("bold"); 
    return rowIndex + 1;
  }

  function addTypeRowWithEmbeddedTotals(sheet, type, rowIndex, columnIndices) {
    const headers = config.getSection('HEADERS');
    const sheetNames = config.getSection('SHEETS');
    const uiColors = config.getSection('COLORS').UI;
    const style = {
        backgroundColor: uiColors.TYPE_HEADER_TOTAL_BG || "#f0f0f0", 
        fontColor: uiColors.TYPE_HEADER_TOTAL_FONT || "#000000",     
        labelFontWeight: "bold",
        numberFontWeight: "normal"
    };
    sheet.getRange(rowIndex, 1).setValue(type).setFontWeight(style.labelFontWeight);
    const fullRowRange = sheet.getRange(rowIndex, 1, 1, headers.length);
    fullRowRange.setBackground(style.backgroundColor).setFontColor(style.fontColor);
    sheet.getRange(rowIndex, 1).setFontWeight(style.labelFontWeight);

    const formulas = [];
    for (let monthCol = 5; monthCol <= 16; monthCol++) {
      const monthDate = new Date(2024, monthCol - 5, 1);
      const formulaParams = { type, monthDate, sheetName: sheetNames.TRANSACTIONS, ...columnIndices, typeCol: columnIndices.type +1, categoryCol: columnIndices.category + 1, subcategoryCol: columnIndices.subcategory + 1, dateCol: columnIndices.date+1, amountCol: columnIndices.amount+1 };
      formulas.push(buildMonthlySumFormula(formulaParams, rowIndex));
    }
    formulas.push(`=SUM(${utils.columnToLetter(5)}${rowIndex}:${utils.columnToLetter(16)}${rowIndex})`);
    formulas.push(`=AVERAGE(${utils.columnToLetter(5)}${rowIndex}:${utils.columnToLetter(16)}${rowIndex})`);
    
    const formulaRange = sheet.getRange(rowIndex, 5, 1, 14);
    formulaRange.setFormulas([formulas]).setFontWeight(style.numberFontWeight).setFontColor(style.fontColor);
    formatRangeAsCurrency(formulaRange, true); 
    return rowIndex + 1;
  }
  
  function addCategoryRows(sheet, combinations, rowIndex, type, columnIndices) {
    if (combinations.length === 0) return rowIndex;
    const expenseTypes = config.getSection('EXPENSE_TYPES');
    const colors = config.getSection('COLORS').UI;
    const sheetNames = config.getSection('SHEETS');
    const startRow = rowIndex; const numRows = combinations.length;
    const categoryData = Array(numRows).fill(null).map(() => Array(3).fill(""));
    combinations.forEach((combo, index) => {
      categoryData[index][0] = combo.type; categoryData[index][1] = combo.category; categoryData[index][2] = combo.subcategory;
    });
    sheet.getRange(startRow, 1, numRows, 3).setValues(categoryData);
    if (expenseTypes.includes(type)) sheet.getRange(startRow, 4, numRows, 1).insertCheckboxes();
    for (let i = 0; i < numRows; i++) {
      const combo = combinations[i]; const currentRow = startRow + i; const rowFormulas = [];
      for (let monthCol = 5; monthCol <= 16; monthCol++) {
        const monthDate = new Date(2024, monthCol - 5, 1);
        const formulaParams = { ...combo, monthDate, sheetName: sheetNames.TRANSACTIONS, ...columnIndices, typeCol: columnIndices.type +1, categoryCol: columnIndices.category + 1, subcategoryCol: columnIndices.subcategory + 1, dateCol: columnIndices.date+1, amountCol: columnIndices.amount+1, sharedCol: columnIndices.shared+1 };
        rowFormulas.push(buildMonthlySumFormula(formulaParams, currentRow));
      }
      if (config.getSection('PERFORMANCE').USE_BATCH_OPERATIONS) sheet.getRange(currentRow, 5, 1, 12).setFormulas([rowFormulas]);
      else for (let monthCol = 0; monthCol < 12; monthCol++) sheet.getRange(currentRow, monthCol + 5).setFormula(rowFormulas[monthCol]);
      sheet.getRange(currentRow, 17).setFormula(`=SUM(E${currentRow}:P${currentRow})`);
      sheet.getRange(currentRow, 18).setFormula(`=AVERAGE(E${currentRow}:P${currentRow})`);
      if (combo.subcategory) sheet.getRange(currentRow, 3).setIndent(5);
      else sheet.getRange(currentRow, 2).setFontWeight("bold");
    }
    const valueRange = sheet.getRange(startRow, 5, numRows, 13); formatRangeAsCurrency(valueRange, false);
    const averageRange = sheet.getRange(startRow, 18, numRows, 1); formatRangeAsCurrency(averageRange, false);
    if (type === config.getSection('TRANSACTION_TYPES').INCOME) {
      for (let i = 0; i < numRows; i++) {
        for (let col = 5; col <= 18; col++) sheet.getRange(startRow + i, col).setFontColor(colors.INCOME_FONT);
      }
    }
    return rowIndex + numRows;
  }

  function addOverallExpensesTotalRow(sheet, rowIndex, expenseTypeTotalRowIndices) {
    const headers = config.getSection('HEADERS');
    const uiColors = config.getSection('COLORS').UI;
    const style = { backgroundColor: uiColors.GRAND_TOTAL_BG || "#bfbfbf", fontColor: uiColors.GRAND_TOTAL_FONT || "#000000", fontWeight: "bold", fontSize: 11 };
    sheet.getRange(rowIndex, 1).setValue("Total Expenses");
    const rowRange = sheet.getRange(rowIndex, 1, 1, headers.length);
    rowRange.setBackground(style.backgroundColor).setFontWeight(style.fontWeight).setFontColor(style.fontColor).setFontSize(style.fontSize).setVerticalAlignment("middle");
    const formulas = [];
    if (expenseTypeTotalRowIndices.length === 0) for (let i = 0; i < 14; i++) formulas.push(0);
    else for (let monthCol = 5; monthCol <= 18; monthCol++) {
      const colLetter = utils.columnToLetter(monthCol);
      const sumParts = expenseTypeTotalRowIndices.map(rn => `${colLetter}${rn}`);
      formulas.push(`=SUM(${sumParts.join(",")})`);
    }
    const formulaRange = sheet.getRange(rowIndex, 5, 1, 14);
    formulaRange.setFormulas([formulas]).setFontColor(style.fontColor);
    formatRangeAsCurrency(formulaRange, true);
    return rowIndex + 1;
  }
  
  function addTypeSubtotalRow(sheet, type, rowIndex, rowCount) { // Used for Income total
    const typeColors = getTypeColors(type); const headers = config.getSection('HEADERS');
    sheet.getRange(rowIndex, 1).setValue(`Total ${type}`);
    const subtotalRowRange = sheet.getRange(rowIndex, 1, 1, headers.length);
    subtotalRowRange.setBackground(typeColors.BG).setFontWeight("bold").setFontColor(typeColors.FONT);
    const formulas = []; const startRowForSum = rowIndex - rowCount; const endRowForSum = rowIndex - 1;
    for (let monthCol = 5; monthCol <= 18; monthCol++) formulas.push(`=SUM(${utils.columnToLetter(monthCol)}${startRowForSum}:${utils.columnToLetter(monthCol)}${endRowForSum})`);
    const formulaRange = sheet.getRange(rowIndex, 5, 1, 14); 
    formulaRange.setFormulas([formulas]).setFontColor(typeColors.FONT);
    formatRangeAsCurrency(formulaRange, true); 
    return rowIndex + 1;
  }
  
  function addNetCalculations(sheet, startRow) {
    const data = sheet.getDataRange().getValues(); const headers = config.getSection('HEADERS');
    const uiColors = config.getSection('COLORS').UI; const totals = findTotalRows(data);
    const { incomeRow, expensesRow, savingsRow } = totals;
    if (!incomeRow || !expensesRow) return startRow;
    sheet.getRange(startRow, 1).setValue("Net Calculations");
    sheet.getRange(startRow, 1, 1, headers.length).setBackground(uiColors.NET_BG).setFontWeight("bold").setFontColor(uiColors.NET_FONT);
    startRow++;
    sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses)");
    sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
    let formulasArr = [];
    for (let col = 5; col <= 18; col++) formulasArr.push(`=${utils.columnToLetter(col)}${incomeRow}-${utils.columnToLetter(col)}${expensesRow}`);
    let numericRange = sheet.getRange(startRow, 5, 1, 14);
    numericRange.setFormulas([formulasArr]).setFontColor("#000000"); formatRangeAsCurrency(numericRange, true);
    startRow++;
    if (savingsRow) {
      sheet.getRange(startRow, 1).setValue("Total Expenses + Savings");
      sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
      formulasArr = [];
      for (let col = 5; col <= 18; col++) formulasArr.push(`=${utils.columnToLetter(col)}${expensesRow}+${utils.columnToLetter(col)}${savingsRow}`);
      numericRange = sheet.getRange(startRow, 5, 1, 14);
      numericRange.setFormulas([formulasArr]).setFontColor("#000000"); formatRangeAsCurrency(numericRange, true);
      startRow++;
      sheet.getRange(startRow, 1).setValue("Net (Total Income - Expenses - Savings)");
      sheet.getRange(startRow, 1, 1, 4).setBackground("#F5F5F5").setFontWeight("bold");
      formulasArr = [];
      for (let col = 5; col <= 18; col++) formulasArr.push(`=${utils.columnToLetter(col)}${incomeRow}-${utils.columnToLetter(col)}${expensesRow}-${utils.columnToLetter(col)}${savingsRow}`);
      numericRange = sheet.getRange(startRow, 5, 1, 14);
      numericRange.setFormulas([formulasArr]).setFontColor("#000000"); formatRangeAsCurrency(numericRange, true);
    }
    sheet.getRange(startRow, 1, 1, headers.length).setBorder(null, null, true, null, null, null, uiColors.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    return startRow + 2;
  }
  
  function findTotalRows(data) {
    const totals = { incomeRow: null, expensesRow: null, savingsRow: null };
    const transactionTypes = config.getSection('TRANSACTION_TYPES');
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === "Total Income") totals.incomeRow = i + 1;
      if (data[i][0] === "Total Expenses") totals.expensesRow = i + 1;
      if (data[i][0] === transactionTypes.SAVINGS) totals.savingsRow = i + 1; // Find "Savings" row for its totals
    }
    return totals;
  }
  
  function formatOverviewSheet(sheet) {
    const headers = config.getSection('HEADERS'); const uiConfig = config.getSection('UI').COLUMN_WIDTHS;
    const colors = config.getSection('COLORS').UI;
    sheet.setColumnWidth(1, uiConfig.TYPE); sheet.setColumnWidth(2, uiConfig.CATEGORY);
    sheet.setColumnWidth(3, uiConfig.SUBCATEGORY); sheet.setColumnWidth(4, uiConfig.SHARED);
    for (let i = 5; i <= 16; i++) sheet.setColumnWidth(i, uiConfig.MONTH);
    sheet.setColumnWidth(17, uiConfig.AVERAGE); sheet.setColumnWidth(18, uiConfig.AVERAGE); 
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && (data[i][0].toString().startsWith("Total ") || data[i][0] === config.getSection('TRANSACTION_TYPES').SAVINGS)) { // Include "Savings" row for border
        sheet.getRange(i + 1, 1, 1, headers.length).setBorder(null, null, true, null, null, null, colors.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
  
  function getTypeColors(type) {
    const typeHeaders = config.getSection('COLORS').TYPE_HEADERS;
    const transactionTypes = config.getSection('TRANSACTION_TYPES');
    let colors = typeHeaders.DEFAULT; const normalizedType = type.toLowerCase();
    for (const key in transactionTypes) {
      if (transactionTypes[key].toLowerCase() === normalizedType && typeHeaders[key]) { colors = typeHeaders[key]; break; }
    }
    if (normalizedType === "wants/pleasure" && typeHeaders.WANTS_PLEASURE) colors = typeHeaders.WANTS_PLEASURE;
    return colors;
  }
  
  function formatRangeAsCurrency(range, isTotalRow = false) {
    const localeConfig = config.getSection('LOCALE');
    const numberFormatString = isTotalRow ? localeConfig.NUMBER_FORMATS.CURRENCY_TOTAL_ROW : localeConfig.NUMBER_FORMATS.CURRENCY_DEFAULT;
    utils.formatAsCurrency(range, numberFormatString);
  }
  
  function getUserPreference(key, defaultValue) {
    try { return settingsService.getValue(key, defaultValue); }
    catch (error) { errorService.log(errorService.create(`Failed to get user preference '${key}'`, { originalError: error.toString(), severity: "medium" })); return defaultValue; }
  }
  
  function setUserPreference(key, value) {
    try { settingsService.setValue(key, value); }
    catch (error) { errorService.log(errorService.create(`Failed to set user preference '${key}'`, { originalError: error.toString(), valueToSet: value, severity: "medium" })); }
  }
  
  class FinancialOverviewBuilder {
    constructor() {
      this.spreadsheet = null; this.overviewSheet = null; this.transactionSheet = null;
      this.showSubCategories = true; this.transactionData = null; this.columnIndices = null;
      this.categoryCombinations = null; this.groupedCombinations = null; this.lastContentRowIndex = 0;
    }
    
    initialize() {
      this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); const sheetNames = config.getSection('SHEETS');
      this.overviewSheet = utils.getOrCreateSheet(this.spreadsheet, sheetNames.OVERVIEW);
      clearSheetContent(this.overviewSheet);
      this.transactionSheet = this.spreadsheet.getSheetByName(sheetNames.TRANSACTIONS);
      if (!this.transactionSheet) throw errorService.create(`Required sheet "${sheetNames.TRANSACTIONS}" not found`, { severity: "high" });
      this.showSubCategories = getUserPreference("ShowSubCategories", true); return this;
    }
    
    processData() {
      const { data, indices } = getProcessedTransactionData(this.transactionSheet);
      this.transactionData = data; this.columnIndices = indices;
      this.categoryCombinations = cacheService.get(config.getSection('CACHE').KEYS.CATEGORY_COMBINATIONS,
        () => getUniqueCategoryCombinations(this.transactionData, this.columnIndices.type, this.columnIndices.category, this.columnIndices.subcategory, this.showSubCategories));
      this.groupedCombinations = cacheService.get(config.getSection('CACHE').KEYS.GROUPED_COMBINATIONS,
        () => groupCategoryCombinations(this.categoryCombinations));
      return this;
    }
    
    setupHeader() { setupHeaderRow(this.overviewSheet, this.showSubCategories); return this; }
    
    generateContent() {
      let rowIndex = 2; 
      const transactionTypes = config.getSection('TRANSACTION_TYPES'); 
      const expenseTypeKeys = config.getSection('EXPENSE_TYPES'); 
      
      const incomeTypeString = transactionTypes.INCOME;
      if (incomeTypeString && this.groupedCombinations[incomeTypeString]) {
        rowIndex = addMajorSectionHeader(this.overviewSheet, incomeTypeString, rowIndex);
        rowIndex = _addSimpleTypeLabelRow(this.overviewSheet, incomeTypeString, rowIndex);
        rowIndex = addCategoryRows(this.overviewSheet, this.groupedCombinations[incomeTypeString], rowIndex, incomeTypeString, this.columnIndices);
        rowIndex = addTypeSubtotalRow(this.overviewSheet, incomeTypeString, rowIndex, this.groupedCombinations[incomeTypeString].length);
        rowIndex += 1; 
      }

      const expenseTypeTotalRowIndices = []; 
      const relevantExpenseTypes = expenseTypeKeys.filter(type => this.groupedCombinations[type] && this.groupedCombinations[type].length > 0);
      if (relevantExpenseTypes.length > 0) {
        rowIndex = addMajorSectionHeader(this.overviewSheet, "Expenses", rowIndex);
        expenseTypeKeys.forEach(typeKey => { 
          const expenseTypeString = typeKey; 
          if (this.groupedCombinations[expenseTypeString] && this.groupedCombinations[expenseTypeString].length > 0) {
            const typeTotalRowIndex = rowIndex;
            rowIndex = addTypeRowWithEmbeddedTotals(this.overviewSheet, expenseTypeString, rowIndex, this.columnIndices);
            expenseTypeTotalRowIndices.push(typeTotalRowIndex); 
            rowIndex = addCategoryRows(this.overviewSheet, this.groupedCombinations[expenseTypeString], rowIndex, expenseTypeString, this.columnIndices);
            rowIndex += 1; 
          }
        });
        if (expenseTypeTotalRowIndices.length > 0) {
          rowIndex = addOverallExpensesTotalRow(this.overviewSheet, rowIndex, expenseTypeTotalRowIndices);
          rowIndex += 1; 
        }
      }

      const savingsTypeString = transactionTypes.SAVINGS; 
      if (savingsTypeString && this.groupedCombinations[savingsTypeString]) {
        rowIndex = addMajorSectionHeader(this.overviewSheet, savingsTypeString, rowIndex);
        rowIndex = addTypeRowWithEmbeddedTotals(this.overviewSheet, savingsTypeString, rowIndex, this.columnIndices); 
        rowIndex = addCategoryRows(this.overviewSheet, this.groupedCombinations[savingsTypeString], rowIndex, savingsTypeString, this.columnIndices);
        rowIndex += 1; 
      }
      
      this.lastContentRowIndex = rowIndex; return this;
    }
    
    addNetCalculations() { this.lastContentRowIndex = addNetCalculations(this.overviewSheet, this.lastContentRowIndex); return this; }
    addMetrics() {
      if (analysisServiceInstance && analysisServiceInstance.analyze) analysisServiceInstance.analyze(this.spreadsheet, this.overviewSheet);
      else { console.error("FinancialAnalysisService not available for addMetrics"); if (errorService) errorService.log(errorService.create("FinancialAnalysisService not available in FinanceOverview", { severity: "high"}));}
      return this;
    }
    formatSheet() { formatOverviewSheet(this.overviewSheet); return this; }
    applyPreferences() { if (this.showSubCategories) this.overviewSheet.showColumns(3, 1); else this.overviewSheet.hideColumns(3, 1); return this; }
    build() { return { sheet: this.overviewSheet, lastRow: this.overviewSheet.getLastRow(), success: true }; }
  }
  
  return {
    create: function() {
      try {
        uiService.showLoadingSpinner("Generating financial overview...");
        cacheService.invalidateAll(); 
        const builder = new FinancialOverviewBuilder();
        const result = builder.initialize().processData().setupHeader().generateContent()
          .addNetCalculations().addMetrics().formatSheet().applyPreferences().build();
        uiService.hideLoadingSpinner();
        uiService.showSuccessNotification("Financial overview generated successfully!");
        return result;
      } catch (error) {
        uiService.hideLoadingSpinner();
        if (error.name === 'FinancialPlannerError') {
          errorService.log(error); uiService.showErrorNotification("Error generating overview", error.message);
        } else {
          const wrappedError = errorService.create("Failed to generate financial overview", { originalError: error.message, stack: error.stack, severity: "high" });
          errorService.log(wrappedError); uiService.showErrorNotification("Error generating overview", error.message);
        }
        throw error;
      }
    },
    handleEdit: function(e) {
      try {
        if (e.range.getSheet().getName() !== config.getSection('SHEETS').OVERVIEW) return;
        const subcategoryToggle = config.getSection('UI').SUBCATEGORY_TOGGLE;
        if (e.range.getA1Notation() === subcategoryToggle.CHECKBOX_CELL) {
          const newValue = e.range.getValue(); 
          setUserPreference("ShowSubCategories", newValue);
          uiService.showLoadingSpinner("Updating overview based on preference change...");
          this.create(); 
        }
      } catch (error) {
         errorService.handle(errorService.create("Error handling Overview sheet edit", { originalError: error.toString(), eventDetails: JSON.stringify(e) }), "Failed to process change on Overview sheet.");
      }
    }
  };
})(
  FinancialPlanner.Utils, FinancialPlanner.UIService, FinancialPlanner.CacheService, FinancialPlanner.ErrorService, 
  FinancialPlanner.Config, FinancialPlanner.SettingsService, FinancialPlanner.FinancialAnalysisService 
);

// ============================================================================
// BACKWARD COMPATIBILITY LAYER
// ============================================================================
function createFinancialOverview() {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.FinanceOverview && FinancialPlanner.FinanceOverview.create) {
    return FinancialPlanner.FinanceOverview.create();
  }
  Logger.log("Global createFinancialOverview: FinancialPlanner.FinanceOverview not available.");
}
function handleOverviewSheetEdits(e) {
  if (typeof FinancialPlanner !== 'undefined' && FinancialPlanner.FinanceOverview && FinancialPlanner.FinanceOverview.handleEdit) {
    FinancialPlanner.FinanceOverview.handleEdit(e);
  } else {
    Logger.log("Global handleOverviewSheetEdits: FinancialPlanner.FinanceOverview not available.");
  }
}
