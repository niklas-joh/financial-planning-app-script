/**
 * @fileoverview Pure helpers that say which bank, connection and account a row
 * belongs to. Labels are joined from the accounts list, never recomputed per row.
 */
(function (ns) {
  /**
   * @typedef {Object} AccountLabels
   * @property {string} bank - Bank name.
   * @property {string} connection - Connection key (bank plus holder).
   * @property {string} accountName - The name set by the user or bank, else the product name, else ''.
   * @property {string} ibanMasked - IBAN reduced to its last four characters, or ''.
   */

  /**
   * Builds a lookup from account id to its labels.
   * @param {Array<import('../providers/bank-provider.js').Account>} accounts - Accounts.
   * @returns {function(string): ?AccountLabels} Labels for a known account id, else null.
   */
  function lookup(accounts) {
    const byId = {};
    accounts.forEach(function (a) {
      byId[a.accountId] = {
        bank: a.bank || '',
        connection: a.connection || '',
        accountName: a.accountName || a.product || '',
        ibanMasked: a.ibanMasked || ''
      };
    });
    return function (accountId) { return byId[accountId] || null; };
  }

  /**
   * Copies a row with fresh labels. Rows of unknown accounts are returned as they are,
   * so another connection's labels are never blanked.
   * @param {Object} row - Row with an `accountId`.
   * @param {function(string): ?AccountLabels} labelsFor - Lookup from `lookup`.
   * @returns {Object} Labelled row.
   */
  function stamp(row, labelsFor) {
    const labels = labelsFor(row.accountId);
    return labels ? Object.assign({}, row, labels) : row;
  }

  ns.AccountLabels = { lookup: lookup, stamp: stamp };
})(globalThis.BankSync);
