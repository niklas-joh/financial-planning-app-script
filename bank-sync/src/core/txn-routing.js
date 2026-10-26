/**
 * @fileoverview Which tab a transaction is stored on, and one store over the three transaction tabs.
 * Rule (owner decisions B1, 2026-10-26): `transactions` holds current-account cash flow, card spending and
 * bank transfers to or from Trading 212 (its cash is used like a current account, so both sides of a
 * transfer are there, for B2 to pair); `investment_transactions` holds the other movements of investment
 * accounts (interest, dividends, fees; every Avanza row); `unclassified_transactions`
 * holds the Trading 212 rows the API cannot tell apart (a WITHDRAW or DEPOSIT may be a card payment, cashback
 * or a bank transfer) until a Trading 212 report labels them. The routed store loads all three tabs as one
 * list and writes each row to its tab on save, so the sync, the merge, the report details and the EUR
 * conversion see one set of rows: a later re-read finds a row wherever it was moved and never adds it again.
 */
(function (ns) {
  /** @returns {Array<string>} The three tabs, in load order. */
  function tabs() {
    const t = ns.Config.TABS;
    return [t.TRANSACTIONS, t.INVESTMENT_TRANSACTIONS, t.UNCLASSIFIED_TRANSACTIONS];
  }

  /**
   * Whether a CSV row comes from an investment account's export. CSV account ids are `<source>-<hash>`
   * (`csv-normalize.js`), so the source is the part before the last hyphen.
   * @param {import('../providers/bank-provider.js').Transaction} t - Row.
   * @returns {boolean} True for a CSV row of a source in `Config.ROUTING.INVESTMENT_CSV_SOURCES`.
   */
  function isInvestmentCsv(t) {
    if (t.provider !== ns.Config.CSV.PROVIDER_ID) return false;
    const id = String(t.accountId);
    const cut = id.lastIndexOf('-');
    return cut > 0 && ns.Config.ROUTING.INVESTMENT_CSV_SOURCES.indexOf(id.slice(0, cut)) !== -1;
  }

  /**
   * The tab a row belongs on. A row the report labelled follows its label (card spending and bank
   * transfers to `transactions`, other movements to `investment_transactions`); otherwise rows of an
   * investment provider or an investment CSV source (Avanza) go to `investment_transactions`, except the
   * movement types that stay unclear without a report; every other row goes to `transactions`.
   * @param {import('../providers/bank-provider.js').Transaction} t - Row (`classification` set by a report, or empty).
   * @returns {string} Tab name.
   */
  function tabFor(t) {
    const names = ns.Config.TABS;
    if (t.classification === 'card' || t.classification === 'transfer') return names.TRANSACTIONS;
    if (t.classification === 'investment') return names.INVESTMENT_TRANSACTIONS;
    if (isInvestmentCsv(t)) return names.INVESTMENT_TRANSACTIONS;
    const providers = ns.Config.ROUTING.INVESTMENT_PROVIDERS;
    if (!Object.prototype.hasOwnProperty.call(providers, t.provider)) return names.TRANSACTIONS;
    return providers[t.provider].unclearTypes.indexOf(t.description) !== -1
      ? names.UNCLASSIFIED_TRANSACTIONS
      : names.INVESTMENT_TRANSACTIONS;
  }

  /**
   * @param {*} value - Cell value.
   * @returns {boolean} True for null, undefined and ''.
   */
  function isEmpty(value) {
    return value === null || value === undefined || value === '';
  }

  /**
   * Joins two stored copies of the same row (left in two tabs when a save stopped halfway): the first
   * copy, with its empty fields taken from the second, so no detail or EUR amount is lost.
   * @param {Object} a - First copy (earlier tab).
   * @param {Object} b - Second copy.
   * @returns {Object} Joined row.
   */
  function join(a, b) {
    const out = Object.assign({}, a);
    Object.keys(b).forEach(function (k) { if (isEmpty(out[k]) && !isEmpty(b[k])) out[k] = b[k]; });
    return out;
  }

  /**
   * Splits rows by the tab they belong on, keeping their order.
   * @param {Array<Object>} rows - Rows.
   * @param {Array<string>} names - Tabs.
   * @returns {Object<string,Array<Object>>} Rows per tab (every tab present).
   */
  function partition(rows, names) {
    const out = {};
    names.forEach(function (tab) { out[tab] = []; });
    rows.forEach(function (t) { out[tabFor(t)].push(t); });
    return out;
  }

  /**
   * @param {Array<Object>} rows - Rows of one tab.
   * @returns {string} What the tab holds, cell for cell, to skip writing a tab that did not change.
   */
  function fingerprint(rows) {
    return JSON.stringify(ns.TransactionsStore.toMatrix(rows));
  }

  /**
   * Creates one store over the three tabs, with the same `load`/`save` as a single tab's store.
   * `save` writes only the tabs whose content changed. A row that changes tab is first written to its
   * new tab while the old tab still has it, and removed from the old tab after that, so a save that
   * stops halfway leaves a row twice (joined on the next load), never nowhere.
   * @param {Object<string,{load: function(): Array<Object>, save: function(Array<Object>): number}>} stores -
   *   One store per tab name (all three tabs).
   * @returns {{load: function(): Array<Object>, save: function(Array<Object>): number,
   *   moved: function(): Object<string,number>}} Store; `moved` counts the rows moved into each tab so far.
   */
  function createStore(stores) {
    const names = tabs();
    const moved = {};
    names.forEach(function (tab) { moved[tab] = 0; });
    /** Rows per tab as last read or written; null until the first load. */
    let onTab = null;
    /** Tabs holding each key as last read or written. */
    let where = {};
    /** Fingerprint per tab as last read or written. */
    let written = {};

    /**
     * Remembers what each tab holds now.
     * @param {Object<string,Array<Object>>} byTab - Rows per tab.
     */
    function remember(byTab) {
      onTab = byTab;
      where = {};
      written = {};
      names.forEach(function (tab) {
        written[tab] = fingerprint(byTab[tab]);
        byTab[tab].forEach(function (t) {
          const key = ns.TxnMerge.rowKey(t);
          (where[key] = where[key] || []).push(tab);
        });
      });
    }

    /**
     * Reads all three tabs. A row found in two tabs is returned once (joined).
     * @returns {Array<import('../providers/bank-provider.js').Transaction>} Rows of every tab.
     */
    function load() {
      const byTab = {};
      names.forEach(function (tab) { byTab[tab] = stores[tab].load(); });
      remember(byTab);
      const byKey = {};
      const order = [];
      names.forEach(function (tab) {
        byTab[tab].forEach(function (t) {
          const key = ns.TxnMerge.rowKey(t);
          if (byKey[key]) byKey[key] = join(byKey[key], t);
          else { byKey[key] = t; order.push(key); }
        });
      });
      return order.map(function (key) { return byKey[key]; });
    }

    /**
     * Writes one tab if its content differs from what it holds.
     * @param {string} tab - Tab.
     * @param {Array<Object>} rows - Its new rows.
     */
    function write(tab, rows) {
      const print = fingerprint(rows);
      if (print === written[tab]) return;
      stores[tab].save(rows);
      written[tab] = print;
    }

    /**
     * Stores the full set of rows, each on its tab.
     * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Every stored row.
     * @returns {number} Rows stored.
     */
    function save(rows) {
      if (!onTab) load();
      const next = partition(rows, names);
      const gaining = {};
      rows.forEach(function (t) {
        const tab = tabFor(t);
        const from = where[ns.TxnMerge.rowKey(t)];
        if (from && from.indexOf(tab) === -1) {
          moved[tab]++;
          gaining[tab] = true;
        }
      });
      // First the tabs that gain rows, still keeping the rows that leave them; then every tab exactly.
      // The second pass cannot lose a row: one that leaves a tab was already written to its new tab by
      // the first pass, so if a write fails in between, the row is on two tabs (joined on load), not none.
      names.forEach(function (tab) {
        if (!gaining[tab]) return;
        const staying = {};
        next[tab].forEach(function (t) { staying[ns.TxnMerge.rowKey(t)] = true; });
        write(tab, next[tab].concat(onTab[tab].filter(function (t) { return !staying[ns.TxnMerge.rowKey(t)]; })));
      });
      names.forEach(function (tab) { write(tab, next[tab]); });
      remember(next);
      return rows.length;
    }

    return {
      load: load,
      save: save,
      moved: function () { return Object.assign({}, moved); }
    };
  }

  /**
   * Text for the rows moved between tabs, e.g. "3 to investment_transactions, 1 to transactions".
   * @param {?Object<string,number>} moved - Counts per tab (from `moved()`), or null.
   * @returns {string} The text, or '' if nothing moved.
   */
  function movedText(moved) {
    if (!moved) return '';
    return tabs().filter(function (tab) { return moved[tab]; })
      .map(function (tab) { return moved[tab] + ' to ' + tab; }).join(', ');
  }

  /**
   * Show status line: rows per transaction tab (counts only).
   * @param {{getSheetByName: function(string): ?{getLastRow: function(): number}}} spreadsheet - Spreadsheet-like.
   * @returns {string} E.g. "Rows: transactions 120, investment_transactions 40, unclassified_transactions 5".
   */
  function statusLine(spreadsheet) {
    return 'Rows: ' + tabs().map(function (tab) {
      const sheet = spreadsheet.getSheetByName(tab);
      return tab + ' ' + (sheet ? Math.max(0, sheet.getLastRow() - 1) : 0);
    }).join(', ');
  }

  ns.TxnRouting = {
    tabs: tabs, tabFor: tabFor, createStore: createStore, movedText: movedText, statusLine: statusLine
  };
})(globalThis.BankSync);
