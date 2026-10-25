/**
 * @fileoverview Provider-agnostic raw refresh: accounts and balances to raw tabs.
 * Reading and writing are separate so several connections can share one write.
 */
(function (ns) {
  /**
   * @typedef {Object} RawData
   * @property {Array<import('../providers/bank-provider.js').Account>} accounts
   * @property {Array<import('../providers/bank-provider.js').Balance>} balances
   * @property {Array<import('../providers/bank-provider.js').Transaction>} [transactions] - What the
   *   latest sync fetched, each with its `apiFields`.
   */

  /**
   * Reads one provider's accounts and their balances.
   * @param {import('../providers/bank-provider.js').BankProvider} provider - Source.
   * @param {Array<import('../providers/bank-provider.js').Account>} accounts - Its accounts.
   * @param {function(): boolean} [outOfTime] - True once the shared time budget is spent.
   * @returns {RawData} Accounts and balances.
   * @throws {Error} Any provider error, e.g. AuthError when consent lapsed; TimeBudget if
   *   `outOfTime` turns true before every account's balances were read (nothing is written then).
   */
  function collect(provider, accounts, outOfTime) {
    const balances = accounts.reduce(function (all, account) {
      if (outOfTime && outOfTime()) throw ns.Errors.create('TimeBudget', 'Time budget reached while reading balances');
      return all.concat(provider.getBalances(account.accountId));
    }, []);
    return { accounts: accounts, balances: balances };
  }

  /**
   * Day a raw row belongs to, as YYYY-MM-DD. Stored cells that Sheets turned into dates come
   * back as full timestamps, so only the date part is used.
   * @param {Object} row - Flat transaction row.
   * @returns {string} ISO date, or '' for an undated pending row.
   */
  function dayOf(row) {
    return ns.TxnMerge.effectiveDate(row).slice(0, 10);
  }

  /**
   * How `raw_transactions` grows: one row per transaction id, newest first, so the raw API
   * fields of earlier fetches are kept even though each sync only reads the last weeks.
   * Rows older than `RAW_RETENTION_DAYS` are dropped so the tab (and each rewrite of it) stays bounded.
   * @param {string} today - ISO date of today.
   * @returns {import('../sheets/raw-writer.js').Accumulate} Policy.
   */
  function accumulateTransactions(today) {
    const oldest = ns.DateRange.addDays(today, -ns.Config.TRANSACTIONS.RAW_RETENTION_DAYS);
    return {
      keyOf: function (t) { return t.transactionId ? ns.TxnMerge.rowKey(t) : ''; },
      keep: function (t) { return !dayOf(t) || dayOf(t) >= oldest; },
      compare: function (a, b) {
        if (dayOf(a) !== dayOf(b)) return dayOf(a) < dayOf(b) ? 1 : -1;
        const ka = ns.TxnMerge.rowKey(a);
        const kb = ns.TxnMerge.rowKey(b);
        return ka < kb ? -1 : ka > kb ? 1 : 0;
      }
    };
  }

  /**
   * Writes the three raw tabs: bank, connection, holder, account name and masked IBAN are added
   * to each balance, and every API field becomes its own `api.*` column.
   * @param {{writeTable: function(string, Array<Object>, Array<Object>, Array<string>=): number}} writer - Sheet writer.
   * @param {RawData} data - Fresh rows from the connections that were read.
   * @param {Array<string>} refreshed - Keys of those connections; other connections' rows stay.
   *   Transactions are the exception: they accumulate across runs (see `accumulateTransactions`).
   * @param {string} today - ISO date of today, for the transactions retention window.
   * @returns {{accounts: number, balances: number, transactions: number}} Row counts written (no sensitive data).
   */
  function write(writer, data, refreshed, today) {
    const labelsFor = ns.AccountLabels.lookup(data.accounts);
    const balances = data.balances.map(function (b) { return ns.RawWriter.toRow(ns.AccountLabels.stamp(b, labelsFor)); });
    const accounts = data.accounts.map(ns.RawWriter.toRow);
    const transactions = (data.transactions || []).map(function (t) { return ns.RawWriter.toRow(ns.AccountLabels.stamp(t, labelsFor)); });
    const tabs = ns.Config.TABS;
    return {
      transactions: writer.writeTable(tabs.RAW_TRANSACTIONS, ns.RawWriter.TRANSACTION_COLUMNS, transactions, refreshed, accumulateTransactions(today)),
      accounts: writer.writeTable(tabs.RAW_ACCOUNTS, ns.RawWriter.ACCOUNT_COLUMNS, accounts, refreshed),
      balances: writer.writeTable(tabs.RAW_BALANCES, ns.RawWriter.BALANCE_COLUMNS, balances, refreshed)
    };
  }

  ns.RawSync = { collect: collect, write: write };
})(globalThis.BankSync);
