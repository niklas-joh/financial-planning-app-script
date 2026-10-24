/**
 * @fileoverview Writes normalized rows to a raw tab in one batch call. Rows of
 * connections that were not refreshed are kept, and the tab is never cleared first.
 * String columns are plain text so bank-supplied strings cannot become formulas.
 */
(function (ns) {
  /**
   * @typedef {Object} Column
   * @property {string} header - Header cell text.
   * @property {string} key - Property read from each row object.
   * @property {boolean} [numeric] - Number column; every other column is plain text.
   */

  /** @type {Array<Column>} */
  const ACCOUNT_COLUMNS = ['provider', 'bank', 'connection', 'accountId', 'accountName', 'currency', 'ibanMasked',
    'accountType', 'product']
    .map(function (k) { return { header: k, key: k }; });

  /** @type {Array<Column>} */
  const BALANCE_COLUMNS = ['provider', 'bank', 'connection', 'accountName', 'ibanMasked', 'accountId',
    'balanceType', 'name', 'amount', 'currency', 'referenceDate', 'lastChange', 'fetchedAt']
    .map(function (k) { return { header: k, key: k, numeric: k === 'amount' }; });

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
   * Reads a tab's existing rows by header name.
   * @param {Object} sheet - Sheet-like.
   * @param {Array<Column>} columns - Column definitions.
   * @returns {Array<Object>} Row objects; [] if the tab is empty.
   */
  function readRows(sheet, columns) {
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    const header = data[0];
    return data.slice(1).map(function (cells) {
      const row = {};
      columns.forEach(function (c) {
        const cell = cells[header.indexOf(c.header)];
        if (cell === undefined) row[c.key] = '';
        else if (cell instanceof Date) row[c.key] = cell.toISOString();
        else if (c.numeric) row[c.key] = cell === '' ? '' : Number(cell);
        else row[c.key] = String(cell);
      });
      return row;
    });
  }

  /**
   * Creates a writer.
   * @param {Object} spreadsheet - Spreadsheet-like (getSheetByName, insertSheet).
   * @returns {{writeTable: function(string, Array<Column>, Array<Object>, Array<string>=): number}} Writer.
   */
  function create(spreadsheet) {
    return {
      /**
       * Writes rows to a tab (created if absent). With `refreshed`, rows of other connections
       * already on the tab are kept (rows without a connection, from before labels existed,
       * are dropped as stale); without it the tab's rows are replaced. The new rows are
       * written first and only leftover rows below are cleared, so a failed write never
       * empties the tab.
       * @param {string} tabName - Target tab.
       * @param {Array<Column>} columns - Column definitions.
       * @param {Array<Object>} rows - Fresh row objects.
       * @param {Array<string>} [refreshed] - Connection keys the fresh rows replace.
       * @returns {number} Number of data rows now on the tab.
       */
      writeTable: function (tabName, columns, rows, refreshed) {
        const sheet = spreadsheet.getSheetByName(tabName) || spreadsheet.insertSheet(tabName);
        const kept = refreshed
          ? readRows(sheet, columns).filter(function (r) { return r.connection && refreshed.indexOf(r.connection) === -1; })
          : [];
        const all = kept.concat(rows);
        const matrix = toMatrix(columns, all);
        const previousLastRow = sheet.getLastRow();
        columns.forEach(function (c, i) {
          if (!c.numeric) sheet.getRange(1, i + 1, matrix.length, 1).setNumberFormat('@');
        });
        sheet.getRange(1, 1, matrix.length, columns.length).setValues(matrix);
        if (previousLastRow > matrix.length) {
          sheet.getRange(matrix.length + 1, 1, previousLastRow - matrix.length, columns.length).clearContent();
        }
        sheet.setFrozenRows(1);
        return all.length;
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
