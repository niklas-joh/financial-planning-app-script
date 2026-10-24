/**
 * Financial Planning Tools - Financial Analysis Service Tests
 *
 * Exercises the real FinancialPlanner.FinancialAnalysisService (and, through
 * it, the real SheetBuilder) against the in-memory Apps Script fakes. The
 * service reads an Overview sheet laid out as columns A–D labels, E–P
 * months, Q total and R average.
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Analysis = FinancialPlanner.FinancialAnalysisService;
  let gas;

  function overviewRow(type, category, sub, monthly) {
    const months = new Array(12).fill(monthly);
    return [type, category, sub, ''].concat(months, [monthly * 12, monthly]);
  }

  function setup() {
    gas = FakeGas.install(global);
    const rows = [
      ['Type', 'Category', 'Sub-Category', 'Shared?'].concat(new Array(12).fill('M'), ['Total', 'Average']),
      overviewRow('Income', 'Salary', '', 4000),
      overviewRow('Total Income', '', '', 4000),
      overviewRow('Essentials', 'Housing', '', -1500),
      overviewRow('Essentials', 'Housing', 'Rent', -1500),
      overviewRow('Essentials', 'Food', '', -500),
      overviewRow('Wants/Pleasure', 'Fun', '', -300),
      overviewRow('Extra', 'Gifts', '', -100),
      overviewRow('Essentials', '', '', -2000),
      overviewRow('Wants/Pleasure', '', '', -300),
      overviewRow('Extra', '', '', -100),
      overviewRow('Total Expenses', '', '', -2400),
      overviewRow('Savings', 'ETF', '', 800),
      overviewRow('Total Savings', '', '', 800),
    ];
    const overview = gas.spreadsheet.insertSheet('Overview');
    overview.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    return overview;
  }

  function analysisSheet() {
    return gas.spreadsheet.getSheetByName(FinancialPlanner.Config.getSheetNames().ANALYSIS);
  }

  function findRow(sheet, label) {
    const values = sheet._values();
    for (let i = 0; i < values.length; i++) if (values[i][0] === label) return { row: i + 1, values: values[i] };
    return null;
  }

  T.registerTest("FinancialAnalysisService", "analyze should create Analysis sheet with title and metric cards", function() {
    const overview = setup();
    Analysis.analyze(gas.spreadsheet, overview);
    const sheet = analysisSheet();
    T.assertNotNull(sheet, "Analysis sheet should be created.");
    T.assertEquals('Financial Analysis', sheet.getRange('A1').getValue(), "Title should be written.");
    T.assertEquals(1, sheet.getFrozenRows(), "Title row should be frozen.");
    const flat = sheet._values().map((r) => r.join('|')).join('\n');
    ['Key Financial Metrics', 'Total Gross Income', 'Overall Savings Rate', 'Final Net Surplus/Deficit'].forEach(function(label) {
      T.assertTrue(flat.includes(label), "Analysis should contain '" + label + "'.");
    });
  });

  T.registerTest("FinancialAnalysisService", "metric cards should reference the Overview totals", function() {
    const overview = setup();
    Analysis.analyze(gas.spreadsheet, overview);
    const sheet = analysisSheet();
    const formulas = [];
    sheet._values().forEach(function(r, i) {
      r.forEach(function(_, c) { const f = sheet.getRange(i + 1, c + 1).getFormula(); if (f) formulas.push(f); });
    });
    T.assertTrue(formulas.indexOf("=IFERROR('Overview'!R3, \"N/A\")") !== -1, "Gross income card should point at Total Income's average (R3).");
    T.assertTrue(formulas.indexOf("=IFERROR('Overview'!R14/'Overview'!R3,0)") !== -1, "Savings rate should divide Total Savings by Total Income.");
  });

  T.registerTest("FinancialAnalysisService", "expense table should list main categories, biggest spend first, with targets", function() {
    const overview = setup();
    Analysis.analyze(gas.spreadsheet, overview);
    const sheet = analysisSheet();
    const header = findRow(sheet, 'Category');
    T.assertNotNull(header, "Expense table header should exist.");
    const body = sheet._values().slice(header.row, header.row + 5).map((r) => [r[0], r[1], r[2], r[4]]);
    T.assertDeepEquals([
      ['Housing', 'Essentials', -1500, 0.5],
      ['Food', 'Essentials', -500, 0.5],
      ['Fun', 'Wants/Pleasure', -300, 0.2],
      ['Gifts', 'Extra', -100, 0.1],
      ['Total Expenses', 'All', -2400, 0.8],
    ], body, "Sub-category rows are excluded and the largest expense comes first.");
    const totalRow = header.row + 5;
    T.assertEquals("=IFERROR(C" + totalRow + "/'Overview'!R3,0)", sheet.getRange(totalRow, 4).getFormula(), "% of income should divide by average income.");
    T.assertEquals("=D" + totalRow + "-E" + totalRow, sheet.getRange(totalRow, 6).getFormula(), "Variance should be actual minus target.");
  });

  T.registerTest("FinancialAnalysisService", "analyze should add the expense charts", function() {
    const overview = setup();
    Analysis.analyze(gas.spreadsheet, overview);
    const charts = analysisSheet().getCharts();
    T.assertEquals(2, charts.length, "A pie and a column chart should be added.");
    T.assertEquals('PIE', charts[0].spec.type, "First chart should be the pie.");
    T.assertEquals('COLUMN', charts[1].spec.type, "Second chart should be the column chart.");
  });

  T.registerTest("FinancialAnalysisService", "re-running analyze replaces the previous charts", function() {
    const overview = setup();
    Analysis.analyze(gas.spreadsheet, overview);
    Analysis.analyze(gas.spreadsheet, overview);
    T.assertEquals(2, analysisSheet().getCharts().length, "Charts should not accumulate across runs.");
  });

  T.registerTest("FinancialAnalysisService", "showKeyMetrics should call analyze and activate sheet", function() {
    setup();
    Analysis.showKeyMetrics();
    T.assertEquals(analysisSheet(), gas.spreadsheet.getActiveSheet(), "Analysis sheet should be activated.");
    T.assertEquals('Financial analysis complete.', gas.spreadsheet.toasts.pop().message, "Success toast should be shown.");
  });

  T.registerTest("FinancialAnalysisService", "showKeyMetrics should show error if Overview sheet missing", function() {
    gas = FakeGas.install(global);
    Analysis.showKeyMetrics();
    T.assertEquals('Error: Overview sheet not found. Please generate the financial overview first.', gas.ui.alerts[0].message, "User should be told to generate the overview.");
    T.assertTrue(analysisSheet() === null, "No Analysis sheet should be created.");
  });

  T.registerTest("FinancialAnalysisService", "global showKeyMetrics should delegate to the service", function() {
    setup();
    global.showKeyMetrics();
    T.assertNotNull(analysisSheet(), "Global wrapper should run the analysis.");
  });
})();
