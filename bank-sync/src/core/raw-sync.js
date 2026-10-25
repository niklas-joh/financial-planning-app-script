/**
 * @fileoverview Provider-agnostic raw refresh: accounts and balances to raw tabs.
 * Reading and writing are separate so several connections can share one write.
 */
(function (ns) {
  /**
   * @typedef {Object} RawData
   * @property {Array<import('../providers/bank-provider.js').Account>} accounts
   * @property {Array<import('../providers/bank-provider.js').Balance>} balances
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
   * Writes both raw tabs: bank, connection, holder, account name and masked IBAN are added
   * to each balance, and every API field becomes its own `api.*` column.
   * @param {{writeTable: function(string, Array<Object>, Array<Object>, Array<string>=): number}} writer - Sheet writer.
   * @param {RawData} data - Fresh rows from the connections that were read.
   * @param {Array<string>} refreshed - Keys of those connections; other connections' rows stay.
   * @returns {{accounts: number, balances: number}} Row counts written (no sensitive data).
   */
  function write(writer, data, refreshed) {
    const labelsFor = ns.AccountLabels.lookup(data.accounts);
    const balances = data.balances.map(function (b) { return ns.RawWriter.toRow(ns.AccountLabels.stamp(b, labelsFor)); });
    const accounts = data.accounts.map(ns.RawWriter.toRow);
    const tabs = ns.Config.TABS;
    return {
      accounts: writer.writeTable(tabs.RAW_ACCOUNTS, ns.RawWriter.ACCOUNT_COLUMNS, accounts, refreshed),
      balances: writer.writeTable(tabs.RAW_BALANCES, ns.RawWriter.BALANCE_COLUMNS, balances, refreshed)
    };
  }

  ns.RawSync = { collect: collect, write: write };
})(globalThis.BankSync);
