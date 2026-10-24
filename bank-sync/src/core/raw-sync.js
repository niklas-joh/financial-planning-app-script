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
   * @returns {RawData} Accounts and balances.
   * @throws {Error} Any provider error, e.g. AuthError when consent lapsed.
   */
  function collect(provider, accounts) {
    const balances = accounts.reduce(function (all, account) {
      return all.concat(provider.getBalances(account.accountId));
    }, []);
    return { accounts: accounts, balances: balances };
  }

  /**
   * Writes both raw tabs, adding bank, connection, account name and masked IBAN to
   * each balance.
   * @param {{writeTable: function(string, Array<Object>, Array<Object>): number}} writer - Sheet writer.
   * @param {RawData} data - Everything to show (all connections).
   * @returns {{accounts: number, balances: number}} Row counts written (no sensitive data).
   */
  function write(writer, data) {
    const labelsFor = ns.AccountLabels.lookup(data.accounts);
    const balances = data.balances.map(function (b) { return ns.AccountLabels.stamp(b, labelsFor); });
    const tabs = ns.Config.TABS;
    return {
      accounts: writer.writeTable(tabs.RAW_ACCOUNTS, ns.RawWriter.ACCOUNT_COLUMNS, data.accounts),
      balances: writer.writeTable(tabs.RAW_BALANCES, ns.RawWriter.BALANCE_COLUMNS, balances)
    };
  }

  ns.RawSync = { collect: collect, write: write };
})(globalThis.BankSync);
