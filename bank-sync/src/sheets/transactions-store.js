/**
 * @fileoverview Reads and writes one transactions tab (`transactions`, `investment_transactions` or
 * `unclassified_transactions`, all with the same columns), one batch call each way.
 * Text columns are forced to plain-text format so bank-supplied strings can never
 * be interpreted as formulas and ids/dates are not reinterpreted by Sheets.
 */
(function (ns) {
  /**
   * @typedef {Object} TxnColumn
   * @property {string} key - Transaction property.
   * @property {boolean} [text] - Stored as plain text.
   * @property {boolean} [nullable] - Empty cell reads back as null.
   * @property {boolean} [optional] - May be missing from an older tab; reads back as ''.
   */

  /** @type {Array<TxnColumn>} */
  const COLUMNS = [
    { key: 'provider', text: true },
    { key: 'bank', text: true, optional: true },
    { key: 'connection', text: true, optional: true },
    { key: 'holderName', text: true, optional: true },
    { key: 'accountName', text: true, optional: true },
    // Name kept for existing sheets and formulas; holds the full IBAN when MASK_ACCOUNT_NUMBERS is off.
    { key: 'ibanMasked', text: true, optional: true },
    { key: 'accountId', text: true },
    { key: 'transactionId', text: true },
    { key: 'bookingDate', text: true },
    { key: 'valueDate', text: true },
    { key: 'status', text: true },
    { key: 'amount' },
    { key: 'currency', text: true },
    { key: 'amountEur', nullable: true },
    { key: 'fxRate', nullable: true },
    { key: 'fxRateDate', text: true },
    { key: 'counterparty', text: true },
    { key: 'description', text: true },
    { key: 'category', text: true, optional: true },
    { key: 'counterpartyAccount', text: true, optional: true },
    { key: 'counterpartyBank', text: true, optional: true },
    { key: 'reference', text: true, optional: true },
    { key: 'classification', text: true, optional: true }
  ];

  /**
   * Converts one cell to the model's value.
   * @param {TxnColumn} column - Column definition.
   * @param {*} cell - Raw cell value (a Date if Sheets parsed a date).
   * @param {function(Date): string} formatDate - Date to ISO date in the sheet timezone.
   * @returns {*} String, number or null.
   */
  function fromCell(column, cell, formatDate) {
    if (cell instanceof Date) return formatDate(cell);
    if (column.nullable) return cell === '' ? null : Number(cell);
    if (column.key === 'amount') return Number(cell);
    return String(cell);
  }

  /**
   * Converts a data matrix to transactions.
   * @param {Array<Array<*>>} matrix - Rows without the header.
   * @param {Object<string,number>} index - Column position of each key (from `indexByHeader`).
   * @param {function(Date): string} formatDate - Date formatter.
   * @returns {Array<import('../providers/bank-provider.js').Transaction>} Transactions.
   */
  function fromMatrix(matrix, index, formatDate) {
    return matrix.map(function (row) {
      const t = {};
      COLUMNS.forEach(function (column) {
        const position = index[column.key];
        t[column.key] = fromCell(column, position === undefined ? '' : row[position], formatDate);
      });
      return t;
    });
  }

  /**
   * Converts transactions to a matrix with a header row.
   * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Transactions.
   * @returns {Array<Array<*>>} Header plus data rows; null becomes ''.
   */
  function toMatrix(rows) {
    const header = COLUMNS.map(function (c) { return c.key; });
    const body = rows.map(function (t) {
      return COLUMNS.map(function (c) { return t[c.key] === null || t[c.key] === undefined ? '' : t[c.key]; });
    });
    return [header].concat(body);
  }

  /**
   * Finds where each column sits, so older tabs without the label columns (or with
   * reordered columns) still load. Fails loudly if a required column is missing.
   * @param {Array<*>} header - First row as read from the sheet.
   * @param {string} tab - Tab name, for the error.
   * @returns {Object<string,number>} Column position by key.
   * @throws {Error} DataError naming the missing required columns.
   */
  function indexByHeader(header, tab) {
    const index = {};
    COLUMNS.forEach(function (column) {
      const position = header.indexOf(column.key);
      if (position !== -1) index[column.key] = position;
    });
    const missing = COLUMNS.filter(function (c) { return !c.optional && index[c.key] === undefined; })
      .map(function (c) { return c.key; });
    if (missing.length) {
      throw ns.Errors.create('DataError', 'The ' + tab +
        ' tab is missing columns (' + missing.join(', ') + '); restore the header or delete the tab to rebuild it');
    }
    return index;
  }

  /**
   * Creates the store.
   * @param {Object} spreadsheet - Spreadsheet-like (getSheetByName, insertSheet).
   * @param {function(Date): string} formatDate - Date to ISO date (YYYY-MM-DD).
   * @param {string} [tabName] - Tab to use (default `transactions`).
   * @returns {{load: function(): Array<Object>, save: function(Array<Object>): number}} Store.
   */
  function create(spreadsheet, formatDate, tabName) {
    const tab = tabName || ns.Config.TABS.TRANSACTIONS;

    /** @returns {Object} The tab, created if absent. */
    function sheet() {
      return spreadsheet.getSheetByName(tab) || spreadsheet.insertSheet(tab);
    }

    return {
      /**
       * Reads every stored transaction with a single range read.
       * @returns {Array<import('../providers/bank-provider.js').Transaction>} Transactions.
       */
      load: function () {
        const existing = spreadsheet.getSheetByName(tab);
        if (!existing || existing.getLastRow() < 1) return [];
        const data = existing.getRange(1, 1, existing.getLastRow(), existing.getLastColumn()).getValues();
        return fromMatrix(data.slice(1), indexByHeader(data[0], tab), formatDate);
      },

      /**
       * Replaces the stored transactions in one write. The new rows are written
       * first and only the leftover rows below them are cleared, so a failed write
       * never empties the tab and cells to the right of the table are untouched.
       * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Full set to store.
       * @returns {number} Rows written.
       */
      save: function (rows) {
        const target = sheet();
        const matrix = toMatrix(rows);
        const previousLastRow = target.getLastRow();
        COLUMNS.forEach(function (column, i) {
          if (column.text) target.getRange(1, i + 1, matrix.length, 1).setNumberFormat('@');
        });
        target.getRange(1, 1, matrix.length, COLUMNS.length).setValues(matrix);
        if (previousLastRow > matrix.length) {
          target.getRange(matrix.length + 1, 1, previousLastRow - matrix.length, COLUMNS.length).clearContent();
        }
        target.setFrozenRows(1);
        return rows.length;
      }
    };
  }

  ns.TransactionsStore = { create: create, COLUMNS: COLUMNS, toMatrix: toMatrix, fromMatrix: fromMatrix };
})(globalThis.BankSync);
