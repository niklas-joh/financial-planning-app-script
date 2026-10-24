/**
 * @fileoverview Reads and writes the `transactions` tab, one batch call each way.
 * Text columns are forced to plain-text format so bank-supplied strings can never
 * be interpreted as formulas and ids/dates are not reinterpreted by Sheets.
 */
(function (ns) {
  /**
   * @typedef {Object} TxnColumn
   * @property {string} key - Transaction property.
   * @property {boolean} [text] - Stored as plain text.
   * @property {boolean} [nullable] - Empty cell reads back as null.
   */

  /** @type {Array<TxnColumn>} */
  const COLUMNS = [
    { key: 'provider', text: true },
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
    { key: 'description', text: true }
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
   * @param {function(Date): string} formatDate - Date formatter.
   * @returns {Array<import('../providers/bank-provider.js').Transaction>} Transactions.
   */
  function fromMatrix(matrix, formatDate) {
    return matrix.map(function (row) {
      const t = {};
      COLUMNS.forEach(function (column, i) { t[column.key] = fromCell(column, row[i], formatDate); });
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
   * Creates the store.
   * @param {Object} spreadsheet - Spreadsheet-like (getSheetByName, insertSheet).
   * @param {function(Date): string} formatDate - Date to ISO date (YYYY-MM-DD).
   * @returns {{load: function(): Array<Object>, save: function(Array<Object>): number}} Store.
   */
  function create(spreadsheet, formatDate) {
    const tab = ns.Config.TABS.TRANSACTIONS;

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
        if (!existing || existing.getLastRow() < 2) return [];
        const data = existing.getRange(2, 1, existing.getLastRow() - 1, COLUMNS.length).getValues();
        return fromMatrix(data, formatDate);
      },

      /**
       * Replaces the tab's contents with the given transactions in one write.
       * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Full set to store.
       * @returns {number} Rows written.
       */
      save: function (rows) {
        const target = sheet();
        const matrix = toMatrix(rows);
        target.clearContents();
        COLUMNS.forEach(function (column, i) {
          if (column.text) target.getRange(1, i + 1, matrix.length, 1).setNumberFormat('@');
        });
        target.getRange(1, 1, matrix.length, COLUMNS.length).setValues(matrix);
        target.setFrozenRows(1);
        return rows.length;
      }
    };
  }

  ns.TransactionsStore = { create: create, COLUMNS: COLUMNS, toMatrix: toMatrix, fromMatrix: fromMatrix };
})(globalThis.BankSync);
