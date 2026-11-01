/**
 * @fileoverview PluxeeProvider: implements the BankProvider contract for one Pluxee card (one
 * connection). One call returns the card and its latest transactions, so it is made once per run and
 * serves `listAccounts` and `getTransactions`; the whole list is read at once (`maxWindowDays: Infinity`)
 * and filtered to the requested dates. No balances yet.
 */
(function (ns) {
  /**
   * @typedef {Object} PluxeeProviderDeps
   * @property {import('../../core/config.js').ConnectionConfig} connection - Connection (holder) served.
   * @property {string} cardId - Card id (`PluxeeCredentials.cardId`).
   * @property {{transactions: function(string): Object}} client - PluxeeClient.
   * @property {function(string): string} toLocalDate - ISO date-time to YYYY-MM-DD in the Bank Sync time zone.
   */

  /**
   * Creates a provider for one connection.
   * @param {PluxeeProviderDeps} deps - Injected services.
   * @returns {import('../bank-provider.js').BankProvider} The provider.
   */
  function create(deps) {
    let response = null;
    /** @returns {{card: Object, transactions: Array<Object>}} The card's response, read once per run. */
    function read() {
      if (!response) response = deps.client.transactions(deps.cardId) || {};
      return response;
    }

    return ns.BankProvider.assertBankProvider({
      maxWindowDays: Infinity,

      authorize: function () {
        throw ns.Errors.config('Pluxee is connected with Script Properties: set ' + ns.Config.PROPERTY_KEYS.PLUXEE_CARD_PREFIX +
          '<HOLDER> and ' + ns.Config.PROPERTY_KEYS.PLUXEE_REFRESH_TOKEN_PREFIX + '<HOLDER> (see the README)');
      },

      listAccounts: function () {
        return [ns.PluxeeNormalize.account(read().card, deps.cardId, deps.connection)];
      },

      getBalances: function () {
        return [];
      },

      /**
       * The card's transactions dated in a range (local dates, inclusive).
       * @param {string} accountId - Card id.
       * @param {{dateFrom: string, dateTo: string}} range - ISO dates.
       * @returns {Array<import('../bank-provider.js').Transaction>} Transactions, declined ones left out.
       */
      getTransactions: function (accountId, range) {
        return ns.PluxeeNormalize.transactions(read().transactions, accountId, deps.toLocalDate).filter(function (t) {
          return t.bookingDate >= range.dateFrom && t.bookingDate <= range.dateTo;
        });
      }
    });
  }

  ns.PluxeeProvider = { create: create };
})(globalThis.BankSync);
