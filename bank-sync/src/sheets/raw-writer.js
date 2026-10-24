/**
 * @fileoverview Writes normalized rows to a tab in one batch call.
 */
(function (ns) {
  /**
   * @typedef {Object} Column
   * @property {string} header - Header cell text.
   * @property {string} key - Property read from each row object.
   */

  /** @type {Array<Column>} */
  const ACCOUNT_COLUMNS = ['provider', 'accountId', 'name', 'currency', 'ibanMasked', 'accountType', 'product']
    .map(function (k) { return { header: k, key: k }; });

  /** @type {Array<Column>} */
  const BALANCE_COLUMNS = ['provider', 'accountId', 'balanceType', 'name', 'amount', 'currency',
    'referenceDate', 'lastChange', 'fetchedAt']
    .map(function (k) { return { header: k, key: k }; });

  /**
   * Builds a 2-D array (header row plus data rows) for `setValues`.
   * @param {Array<Column>} columns - Column definitions.
   * @param {Array<Object>} rows - Row objects.
   * @returns {Array<Array<*>>} Matrix; missing values become ''.
   */
  function toMatrix(columns, rows) {
    const header = columns.map(function (c) { return c.header; });
    const body = rows.map(function (row) {
      return columns.map(function (c) { return row[c.key] === undefined ? '' : row[c.key]; });
    });
    return [header].concat(body);
  }

  /**
   * Creates a writer.
   * @param {Object} spreadsheet - Spreadsheet-like (getSheetByName, insertSheet).
   * @returns {{writeTable: function(string, Array<Column>, Array<Object>): number}} Writer.
   */
  function create(spreadsheet) {
    return {
      /**
       * Replaces a tab's contents with the given rows.
       * @param {string} tabName - Target tab (created if absent).
       * @param {Array<Column>} columns - Column definitions.
       * @param {Array<Object>} rows - Row objects.
       * @returns {number} Number of data rows written.
       */
      writeTable: function (tabName, columns, rows) {
        const sheet = spreadsheet.getSheetByName(tabName) || spreadsheet.insertSheet(tabName);
        const matrix = toMatrix(columns, rows);
        sheet.clearContents();
        sheet.getRange(1, 1, matrix.length, columns.length).setValues(matrix);
        sheet.setFrozenRows(1);
        return rows.length;
      }
    };
  }

  ns.RawWriter = {
    create: create,
    toMatrix: toMatrix,
    ACCOUNT_COLUMNS: ACCOUNT_COLUMNS,
    BALANCE_COLUMNS: BALANCE_COLUMNS
  };
})(globalThis.BankSync);
