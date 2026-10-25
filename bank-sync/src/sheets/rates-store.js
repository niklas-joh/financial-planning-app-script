/**
 * @fileoverview Reads and writes the Rates tab: one row per currency and date, newest
 * first, `rate` = units of `currency` per 1 `base` (ECB reference rate). One batch call
 * each way; writing reuses the raw writer (plain-text dates, never a whole-tab clear).
 */
(function (ns) {
  /** @type {Array<import('./raw-writer.js').Column>} */
  const COLUMNS = ['date', 'base', 'currency', 'rate', 'source']
    .map(function (k) { return { header: k, key: k, numeric: k === 'rate' }; });

  /**
   * Turns tab rows into a rate table, skipping rows that are not usable rates.
   * @param {Array<Array<*>>} data - Header plus data rows.
   * @param {function(Date): string} formatDate - Date cell to ISO date.
   * @returns {import('../fx/fx.js').RateTable} Rates per currency, ascending.
   */
  function fromMatrix(data, formatDate) {
    const header = data[0] || [];
    const at = function (key) { return header.indexOf(key); };
    const table = {};
    data.slice(1).forEach(function (cells) {
      const raw = cells[at('date')];
      const date = raw instanceof Date ? formatDate(raw) : String(raw);
      const currency = String(cells[at('currency')]);
      const rate = Number(cells[at('rate')]);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !currency || !(rate > 0)) return;
      (table[currency] = table[currency] || []).push({ date: date, rate: rate });
    });
    Object.keys(table).forEach(function (currency) { table[currency] = ns.Fx.mergeSeries([], table[currency]); });
    return table;
  }

  /**
   * Turns a rate table into flat rows, by currency, newest first.
   * @param {import('../fx/fx.js').RateTable} table - Rates.
   * @returns {Array<Object>} Rows for the writer.
   */
  function toRows(table) {
    return Object.keys(table).sort().reduce(function (rows, currency) {
      return rows.concat(table[currency].slice().reverse().map(function (r) {
        return { date: r.date, base: ns.Config.REPORTING_CURRENCY, currency: currency, rate: r.rate, source: ns.Config.FX.SOURCE_LABEL };
      }));
    }, []);
  }

  /**
   * Creates the store.
   * @param {Object} spreadsheet - Spreadsheet-like (getSheetByName, insertSheet).
   * @param {function(Date): string} formatDate - Date to ISO date (YYYY-MM-DD).
   * @returns {{load: function(): import('../fx/fx.js').RateTable, save: function(import('../fx/fx.js').RateTable): number}} Store.
   */
  function create(spreadsheet, formatDate) {
    const tab = ns.Config.TABS.RATES;
    return {
      /** @returns {import('../fx/fx.js').RateTable} Stored rates; {} if the tab is absent or empty. */
      load: function () {
        const sheet = spreadsheet.getSheetByName(tab);
        if (!sheet || sheet.getLastRow() < 2) return {};
        return fromMatrix(sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues(), formatDate);
      },
      /**
       * @param {import('../fx/fx.js').RateTable} table - Every rate to keep.
       * @returns {number} Rows written.
       */
      save: function (table) {
        return ns.RawWriter.create(spreadsheet).writeTable(tab, COLUMNS, toRows(table));
      }
    };
  }

  ns.RatesStore = { create: create, COLUMNS: COLUMNS };
})(globalThis.BankSync);
