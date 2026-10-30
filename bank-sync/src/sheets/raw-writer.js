/**
 * @fileoverview Writes normalized rows plus every API field to a raw tab in one batch call.
 * Rows of connections that were not refreshed are kept, and the tab is never cleared first.
 * String columns are plain text so bank-supplied strings cannot become formulas.
 */
(function (ns) {
  /** Row keys with this prefix are API fields; they become extra columns after the fixed ones. */
  const API_PREFIX = 'api.';

  /**
   * @typedef {Object} Column
   * @property {string} header - Header cell text.
   * @property {string} key - Property read from each row object.
   * @property {boolean} [numeric] - Number column; text columns are formatted as plain text.
   */

  /**
   * Columns from keys; the listed keys are numeric.
   * @param {Array<string>} keys - Column keys, in order.
   * @param {Array<string>} numericKeys - Keys holding numbers.
   * @returns {Array<Column>} Columns.
   */
  function columns(keys, numericKeys) {
    return keys.map(function (k) { return { header: k, key: k, numeric: numericKeys.indexOf(k) !== -1 }; });
  }

  // `ibanMasked` keeps its name on every tab so sheets and formulas using it still work; it holds the full IBAN
  // when Config.MASK_ACCOUNT_NUMBERS is off (see Redact.maskAccount).

  /** @type {Array<Column>} */
  const ACCOUNT_COLUMNS = columns(['provider', 'bank', 'connection', 'holderName', 'accountId', 'accountName', 'currency',
    'ibanMasked', 'accountType', 'product', 'detailsAt'], []);

  /** @type {Array<Column>} */
  const BALANCE_COLUMNS = columns(['provider', 'bank', 'connection', 'holderName', 'accountName', 'ibanMasked', 'accountId',
    'balanceType', 'description', 'valueKind', 'amount', 'currency', 'amountEur', 'fxRate', 'fxRateDate', 'referenceDate', 'lastChange', 'fetchedAt'],
  ['amount', 'amountEur', 'fxRate']);

  /** @type {Array<Column>} */
  const TRANSACTION_COLUMNS = columns(['provider', 'bank', 'connection', 'holderName', 'accountName', 'ibanMasked', 'accountId',
    'transactionId', 'status', 'bookingDate', 'valueDate', 'amount', 'currency', 'counterparty', 'description'], ['amount']);

  /** @type {Array<Column>} Open positions; values in the account currency, prices in the instrument's. */
  const HOLDING_COLUMNS = columns(['provider', 'bank', 'connection', 'holderName', 'accountName', 'accountId', 'ticker', 'isin',
    'name', 'instrumentCurrency', 'quantity', 'averagePrice', 'currentPrice', 'amount', 'currency', 'amountEur', 'fxRate',
    'fxRateDate', 'totalCost', 'unrealizedPnl', 'fxImpact', 'openedAt', 'asOf'],
  ['quantity', 'averagePrice', 'currentPrice', 'amount', 'amountEur', 'fxRate', 'totalCost', 'unrealizedPnl', 'fxImpact']);

  /** @type {Array<Column>} One row per account per day (see `core/daily-values.js`). */
  const DAILY_VALUE_COLUMNS = columns(['date', 'provider', 'bank', 'connection', 'holderName', 'accountName', 'accountId',
    'currency', 'cash', 'investments', 'total', 'cashEur', 'investmentsEur', 'totalEur', 'fxRate', 'fxRateDate', 'fetchedAt'],
  ['cash', 'investments', 'total', 'cashEur', 'investmentsEur', 'totalEur', 'fxRate']);

  /**
   * Turns an account, balance or transaction into a flat row: its own fields, plus each API field as `api.<name>`.
   * @param {Object} item - Account, Balance or Transaction, optionally with `apiFields`.
   * @returns {Object} Flat row.
   */
  function toRow(item) {
    const row = Object.assign({}, item);
    delete row.apiFields;
    Object.keys(item.apiFields || {}).forEach(function (key) { row[API_PREFIX + key] = item.apiFields[key]; });
    return row;
  }

  /**
   * Fixed columns followed by one column per distinct API field, sorted by name. A field that
   * is only ever empty because some rows send `null` where others send an object (`creditor`
   * next to `creditor.name`) gets no column of its own.
   * @param {Array<Column>} fixed - Fixed columns.
   * @param {Array<Object>} rows - Flat rows.
   * @returns {Array<Column>} All columns.
   */
  function withApiColumns(fixed, rows) {
    const seen = {};
    rows.forEach(function (row) {
      Object.keys(row).forEach(function (key) { if (key.indexOf(API_PREFIX) === 0) seen[key] = true; });
    });
    const keys = Object.keys(seen).sort();
    const parents = keys.filter(function (key) {
      return keys.some(function (other) { return other.indexOf(key + '.') === 0 || other.indexOf(key + '[') === 0; });
    });
    return fixed.concat(keys.filter(function (key) { return parents.indexOf(key) === -1; })
      .map(function (key) { return { header: key, key: key }; }));
  }

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
   * Reads a tab back into flat rows keyed by header, including the API columns.
   * @param {Object} sheet - Sheet-like.
   * @param {Array<Column>} fixed - Fixed columns (numeric ones are converted back to numbers).
   * @returns {Array<Object>} Rows; [] if the tab is empty.
   */
  function readRows(sheet, fixed) {
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    const numeric = {};
    fixed.forEach(function (c) { if (c.numeric) numeric[c.key] = true; });
    return data.slice(1).map(function (cells) {
      const row = {};
      data[0].forEach(function (header, i) {
        const cell = cells[i];
        if (header === '') return;
        if (cell instanceof Date) row[header] = cell.toISOString();
        else row[header] = numeric[header] && cell !== '' ? Number(cell) : cell;
      });
      return row;
    });
  }

  /**
   * Whether a column holds text (and so should be plain-text formatted).
   * @param {Column} column - Column.
   * @param {Array<Object>} rows - Rows being written.
   * @returns {boolean} True if not numeric and at least one value is a string.
   */
  function isText(column, rows) {
    return !column.numeric && rows.some(function (r) { return typeof r[column.key] === 'string' && r[column.key] !== ''; });
  }

  /**
   * Stored rows of connections that were not refreshed, followed by the fresh rows.
   * @param {Array<Object>} stored - Rows read from the tab.
   * @param {Array<Object>} fresh - Fresh flat rows.
   * @param {Array<string>} [refreshed] - Connection keys the fresh rows replace; none means replace the tab.
   * @returns {Array<Object>} Rows to write.
   */
  function keepOthers(stored, fresh, refreshed) {
    const kept = refreshed
      ? stored.filter(function (r) { return r.connection && refreshed.indexOf(r.connection) === -1; })
      : [];
    return kept.concat(fresh);
  }

  /**
   * @typedef {Object} Accumulate
   * @property {function(Object): string} keyOf - Identity of a row; '' means it has none and is kept as it is.
   * @property {function(Object, Object): number} compare - Order of the written rows.
   * @property {function(Object): boolean} keep - False for rows that have aged out and are dropped.
   * @property {function(Object): Object} [refresh] - Applied to each stored row before merging (e.g. fresh account
   *   labels), so rows no fetch returns again still follow their account.
   */

  /**
   * Stored rows plus fresh ones, one row per key (the fresh copy wins), aged-out rows dropped,
   * in the requested order.
   * @param {Array<Object>} stored - Rows read from the tab.
   * @param {Array<Object>} fresh - Fresh flat rows.
   * @param {Accumulate} how - Identity, retention and order.
   * @returns {Array<Object>} Rows to write.
   */
  function accumulated(stored, fresh, how) {
    const byKey = {};
    const unkeyed = [];
    const kept = how.refresh ? stored.map(how.refresh) : stored;
    kept.concat(fresh).forEach(function (row) {
      const key = how.keyOf(row);
      if (key) byKey[key] = row;
      else unkeyed.push(row);
    });
    return Object.keys(byKey).map(function (key) { return byKey[key]; }).concat(unkeyed).filter(how.keep).sort(how.compare);
  }

  /**
   * Creates a writer.
   * @param {Object} spreadsheet - Spreadsheet-like (getSheetByName, insertSheet).
   * @returns {{writeTable: function(string, Array<Column>, Array<Object>, Array<string>=): number,
   *   readTable: function(string, Array<Column>): Array<Object>}} Writer.
   */
  function create(spreadsheet) {
    return {
      /**
       * Reads a tab back as flat rows keyed by header (API columns included).
       * @param {string} tabName - Tab.
       * @param {Array<Column>} fixed - Its fixed columns (numeric ones come back as numbers).
       * @returns {Array<Object>} Rows; [] if the tab is missing or empty.
       */
      readTable: function (tabName, fixed) {
        const sheet = spreadsheet.getSheetByName(tabName);
        return sheet ? readRows(sheet, fixed) : [];
      },

      /**
       * Writes rows to a tab (created if absent). With `refreshed`, rows of other connections
       * already on the tab are kept (rows without a connection, from before labels existed,
       * are dropped as stale); without it the tab's rows are replaced. The new rows are
       * written first and only leftover cells below and to the right are cleared, so a failed
       * write never empties the tab.
       * @param {string} tabName - Target tab.
       * @param {Array<Column>} fixed - Fixed columns; API columns are added from the rows.
       * @param {Array<Object>} rows - Fresh flat rows (see `toRow`).
       * @param {Array<string>} [refreshed] - Connection keys the fresh rows replace.
       * @param {Accumulate} [accumulate] - Keep rows across runs instead: a fresh row replaces the stored row
       *   with the same key, other stored rows stay until `keep` says they have aged out, and the result is
       *   ordered by `compare`. Replaces the `refreshed` behaviour.
       * @returns {number} Number of data rows now on the tab.
       */
      writeTable: function (tabName, fixed, rows, refreshed, accumulate) {
        const sheet = spreadsheet.getSheetByName(tabName) || spreadsheet.insertSheet(tabName);
        const all = accumulate ? accumulated(readRows(sheet, fixed), rows, accumulate) : keepOthers(readRows(sheet, fixed), rows, refreshed);
        const columns = withApiColumns(fixed, all);
        const matrix = toMatrix(columns, all);
        const previousRows = sheet.getLastRow();
        const previousCols = sheet.getLastColumn();
        columns.forEach(function (c, i) {
          if (isText(c, all)) sheet.getRange(1, i + 1, matrix.length, 1).setNumberFormat('@');
        });
        sheet.getRange(1, 1, matrix.length, columns.length).setValues(matrix);
        if (previousRows > matrix.length) {
          sheet.getRange(matrix.length + 1, 1, previousRows - matrix.length, Math.max(previousCols, columns.length)).clearContent();
        }
        if (previousCols > columns.length) {
          sheet.getRange(1, columns.length + 1, Math.min(previousRows, matrix.length), previousCols - columns.length).clearContent();
        }
        sheet.setFrozenRows(1);
        return all.length;
      }
    };
  }

  ns.RawWriter = {
    create: create,
    toRow: toRow,
    toMatrix: toMatrix,
    ACCOUNT_COLUMNS: ACCOUNT_COLUMNS,
    BALANCE_COLUMNS: BALANCE_COLUMNS,
    HOLDING_COLUMNS: HOLDING_COLUMNS,
    DAILY_VALUE_COLUMNS: DAILY_VALUE_COLUMNS,
    TRANSACTION_COLUMNS: TRANSACTION_COLUMNS,
    API_PREFIX: API_PREFIX
  };
})(globalThis.BankSync);
