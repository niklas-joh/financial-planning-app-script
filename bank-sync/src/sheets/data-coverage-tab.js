/**
 * @fileoverview The `data_coverage` tab (written whole by every Sync all, never typed in) and the hand-kept
 * `manual_balances` tab (only read). Both are read by header name; one batch call each way. The coverage tab is
 * overwritten from the top and only the rows below the new ones are cleared, never the whole tab, so formatting
 * the owner adds stays.
 */
(function (ns) {
  /**
   * @param {Object} sheet - A tab.
   * @param {function(Date): string} formatDate - Date to YYYY-MM-DD.
   * @returns {Array<Object>} Its rows below the header, keyed by header; dates as YYYY-MM-DD.
   */
  function rowsOf(sheet, formatDate) {
    if (!sheet || sheet.getLastRow() < 2) return [];
    const width = sheet.getLastColumn();
    const values = sheet.getRange(1, 1, sheet.getLastRow(), width).getValues();
    const header = values[0].map(function (h) { return String(h).trim(); });
    return values.slice(1).map(function (cells) {
      const row = {};
      header.forEach(function (h, i) {
        if (!h) return;
        const v = cells[i];
        row[h] = v instanceof Date ? formatDate(v) : String(v === null || v === undefined ? '' : v).trim();
      });
      return row;
    });
  }

  /**
   * @param {Object} spreadsheet - Spreadsheet-like (getSheetByName, insertSheet).
   * @param {function(Date): string} formatDate - Date to YYYY-MM-DD in the Bank Sync time zone.
   * @returns {{load: function(): Array<Object>, loadManual: function(): Array<Object>,
   *   write: function(Array<Object>): void}} The coverage tab, and the manual balances ([] when there is no tab).
   */
  function create(spreadsheet, formatDate) {
    const TABS = ns.Config.TABS;
    return {
      load: function () { return rowsOf(spreadsheet.getSheetByName(TABS.DATA_COVERAGE), formatDate); },
      loadManual: function () { return rowsOf(spreadsheet.getSheetByName(TABS.MANUAL_BALANCES), formatDate); },
      write: function (rows) {
        const header = ns.DataCoverage.HEADER;
        const sheet = spreadsheet.getSheetByName(TABS.DATA_COVERAGE) || spreadsheet.insertSheet(TABS.DATA_COVERAGE);
        const before = sheet.getLastRow();
        const matrix = [header].concat(rows.map(function (r) {
          return header.map(function (h) { return r[h] === undefined ? '' : r[h]; });
        }));
        const range = sheet.getRange(1, 1, matrix.length, header.length);
        range.setNumberFormat('@');
        range.setValues(matrix);
        if (before > matrix.length) sheet.getRange(matrix.length + 1, 1, before - matrix.length, header.length).clearContent();
        sheet.setFrozenRows(1);
      }
    };
  }

  ns.DataCoverageTab = { create: create };
})(globalThis.BankSync);
