/**
 * @fileoverview Common provider contract and normalized data models.
 * Providers are independent: adding one never requires editing another.
 */
(function (ns) {
  /**
   * @typedef {Object} Account
   * @property {string} provider - Provider id, e.g. 'enablebanking'.
   * @property {string} accountId - Provider-scoped account identifier.
   * @property {string} name - Display name.
   * @property {string} currency - ISO 4217 code.
   * @property {string} ibanMasked - IBAN reduced to last four characters, or ''.
   * @property {string} accountType - Provider account type (e.g. ISO cash account type).
   * @property {string} product - Provider product name, or ''.
   */

  /**
   * @typedef {Object} Balance
   * @property {string} provider - Provider id.
   * @property {string} accountId - Account the balance belongs to.
   * @property {string} balanceType - Provider balance type code (e.g. CLBD).
   * @property {string} name - Provider balance label.
   * @property {number} amount - Amount in the account currency.
   * @property {string} currency - ISO 4217 code.
   * @property {string} referenceDate - ISO date, or ''.
   * @property {string} lastChange - ISO date-time, or ''.
   * @property {string} fetchedAt - ISO date-time the balance was read.
   */

  /**
   * @typedef {Object} Transaction
   * @property {string} provider
   * @property {string} accountId
   * @property {string} transactionId
   * @property {string} bookingDate
   * @property {string} valueDate
   * @property {string} status - 'pending' or 'booked'.
   * @property {number} amount
   * @property {string} currency
   * @property {?number} amountEur
   * @property {?number} fxRate
   * @property {string} fxRateDate
   * @property {string} counterparty
   * @property {string} description
   */

  /**
   * @typedef {Object} BankProvider
   * @property {function(Object): Object} authorize - Starts user authorization.
   * @property {function(): Array<Account>} listAccounts
   * @property {function(string): Array<Balance>} getBalances
   * @property {function(string, Object): Array<Transaction>} getTransactions
   */

  const REQUIRED_METHODS = ['authorize', 'listAccounts', 'getBalances', 'getTransactions'];

  /**
   * Verifies an object implements the provider contract.
   * @param {Object} candidate - Possible provider.
   * @returns {Object} The same object, if valid.
   * @throws {Error} ConfigError naming the missing methods.
   */
  function assertBankProvider(candidate) {
    const missing = REQUIRED_METHODS.filter(function (name) {
      return !candidate || typeof candidate[name] !== 'function';
    });
    if (missing.length) {
      throw ns.Errors.config('Provider is missing methods: ' + missing.join(', '));
    }
    return candidate;
  }

  ns.BankProvider = { REQUIRED_METHODS: REQUIRED_METHODS, assertBankProvider: assertBankProvider };
})(globalThis.BankSync);
