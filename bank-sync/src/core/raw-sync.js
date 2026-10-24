/**
 * @fileoverview Provider-agnostic raw refresh: accounts and balances to raw tabs.
 */
(function (ns) {
  /**
   * Reads accounts and their balances from a provider and writes both raw tabs.
   * @param {import('../providers/bank-provider.js').BankProvider} provider - Source.
   * @param {{writeTable: function(string, Array<Object>, Array<Object>): number}} writer - Sheet writer.
   * @returns {{accounts: number, balances: number}} Row counts written (no sensitive data).
   * @throws {Error} Any provider error, e.g. AuthError when consent lapsed.
   */
  function run(provider, writer) {
    const accounts = provider.listAccounts();
    const balances = accounts.reduce(function (all, account) {
      return all.concat(provider.getBalances(account.accountId));
    }, []);
    const tabs = ns.Config.TABS;
    return {
      accounts: writer.writeTable(tabs.RAW_ACCOUNTS, ns.RawWriter.ACCOUNT_COLUMNS, accounts),
      balances: writer.writeTable(tabs.RAW_BALANCES, ns.RawWriter.BALANCE_COLUMNS, balances)
    };
  }

  ns.RawSync = { run: run };
})(globalThis.BankSync);
