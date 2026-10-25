/**
 * @fileoverview Pure helpers that say which bank, connection and account a row
 * belongs to. Labels are joined from the accounts list, never recomputed per row.
 */
(function (ns) {
  /**
   * @typedef {Object} AccountLabels
   * @property {string} bank - Bank name.
   * @property {string} connection - Connection key (bank plus holder).
   * @property {string} holderName - Account holder's name as the bank reports it, or ''.
   * @property {string} accountName - The name set by the user or bank, else the product name, else account type and currency (e.g. 'Current account EUR').
   * @property {string} ibanMasked - IBAN reduced to its last four characters, or ''.
   */

  /** Readable names for ISO 20022 cash account type codes; unknown codes are shown as they are. */
  const TYPE_NAMES = { CACC: 'Current account', SVGS: 'Savings', CARD: 'Card', CASH: 'Cash', LOAN: 'Loan', TRAN: 'Transacting account' };

  /**
   * Best display name: the name set by the user or bank, else the product, else the account
   * type and currency (banks such as Revolut send no name), so rows can still be grouped.
   * @param {import('../providers/bank-provider.js').Account} a - Account.
   * @returns {string} Display name, '' only if nothing is known.
   */
  function displayName(a) {
    const type = TYPE_NAMES[a.accountType] || a.accountType || '';
    return a.accountName || a.product || [type, a.currency].filter(Boolean).join(' ');
  }

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
        holderName: a.holderName || '',
        accountName: displayName(a),
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
