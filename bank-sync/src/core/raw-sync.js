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
   * How `raw_transactions` grows: one row per transaction id, newest first, so the raw API
   * fields of earlier fetches are kept even though each sync only reads the last weeks.
   * @type {{keyOf: function(Object): string, compare: function(Object, Object): number}}
   */
  const accumulateTransactions = {
    keyOf: ns.TxnMerge.rowKey,
    compare: function (a, b) {
      const da = ns.TxnMerge.effectiveDate(a);
      const db = ns.TxnMerge.effectiveDate(b);
      if (da !== db) return da < db ? 1 : -1;
      return 0;
    }
  };

  /**
   * Writes the three raw tabs: bank, connection, holder, account name and masked IBAN are added
   * to each balance, and every API field becomes its own `api.*` column.
   * @param {{writeTable: function(string, Array<Object>, Array<Object>, Array<string>=): number}} writer - Sheet writer.
   * @param {RawData} data - Fresh rows from the connections that were read.
   * @param {Array<string>} refreshed - Keys of those connections; other connections' rows stay.
   *   Transactions are the exception: they accumulate across runs (see `accumulateTransactions`).
   * @returns {{accounts: number, balances: number, transactions: number}} Row counts written (no sensitive data).
   */
  function write(writer, data, refreshed) {
    const labelsFor = ns.AccountLabels.lookup(data.accounts);
    const balances = data.balances.map(function (b) { return ns.RawWriter.toRow(ns.AccountLabels.stamp(b, labelsFor)); });
    const accounts = data.accounts.map(ns.RawWriter.toRow);
    const transactions = (data.transactions || []).map(function (t) { return ns.RawWriter.toRow(ns.AccountLabels.stamp(t, labelsFor)); });
    const tabs = ns.Config.TABS;
    return {
      transactions: writer.writeTable(tabs.RAW_TRANSACTIONS, ns.RawWriter.TRANSACTION_COLUMNS, transactions, refreshed, accumulateTransactions),
      accounts: writer.writeTable(tabs.RAW_ACCOUNTS, ns.RawWriter.ACCOUNT_COLUMNS, accounts, refreshed),
      balances: writer.writeTable(tabs.RAW_BALANCES, ns.RawWriter.BALANCE_COLUMNS, balances, refreshed)
    };
  }

  ns.RawSync = { collect: collect, write: write };
})(globalThis.BankSync);
