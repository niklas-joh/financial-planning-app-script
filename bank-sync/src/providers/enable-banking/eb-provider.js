/**
 * @fileoverview EnableBankingProvider: implements the BankProvider contract
 * for one bank over Enable Banking.
 */
(function (ns) {
  /**
   * @typedef {Object} EbProviderDeps
   * @property {import('../../core/config.js').ConnectionConfig} connection - Login this instance serves.
   * @property {Object} client - EbClient.
   * @property {Object} store - EbSessionStore.
   * @property {Object} auth - EbAuth flow.
   * @property {function(): number} nowMs - Clock.
   */

  /**
   * Creates a provider for one connection (one login at one bank).
   * @param {EbProviderDeps} deps - Injected services.
   * @returns {import('../bank-provider.js').BankProvider} The provider.
   */
  function create(deps) {
    /**
     * Loads the stored session and checks it is usable.
     * @returns {import('./eb-session-store.js').StoredSession} Valid session.
     * @throws {Error} AuthError if never authorized, lapsed or revoked.
     */
    function requireSession() {
      const session = deps.store.loadSession(deps.connection.key);
      if (!session) throw ns.Errors.auth('Not authorized yet; use "Connect a bank..." first');
      if (Date.parse(session.accessValidUntil) <= deps.nowMs()) {
        throw ns.Errors.auth('Consent has expired; re-authorize');
      }
      return session;
    }

    /**
     * Adds the bank and connection this provider serves, and the stored API fields.
     * @param {import('../bank-provider.js').Account} account - Stored account.
     * @returns {import('../bank-provider.js').Account} Account with `bank`, `connection` and `apiFields` set.
     */
    function stampAccount(account) {
      return Object.assign({}, account, {
        bank: deps.connection.bank.aspspName,
        connection: deps.connection.key,
        apiFields: deps.store.loadApiFields(account.accountId)
      });
    }

    return ns.BankProvider.assertBankProvider({
      authorize: function () {
        return deps.auth.begin(deps.connection);
      },

      listAccounts: function (options) {
        const session = requireSession();
        const state = deps.client.getSession(session.sessionId);
        if (state.status && state.status !== 'AUTHORIZED') {
          throw ns.Errors.auth('Session status is ' + state.status + '; re-authorize');
        }
        return deps.auth.refreshStale(deps.connection, options && options.outOfTime).map(stampAccount);
      },

      getBalances: function (accountId) {
        requireSession();
        const fetchedAt = new Date(deps.nowMs()).toISOString();
        return (deps.client.getBalances(accountId).balances || []).map(function (raw) {
          return ns.EbNormalize.normalizeBalance(raw, accountId, fetchedAt);
        });
      },

      /**
       * Fetches every page of an account's transactions for a date range.
       * @param {string} accountId - Account uid.
       * @param {{dateFrom: string, dateTo: string}} range - Inclusive ISO dates (UTC).
       * @returns {Array<import('../bank-provider.js').Transaction>} Normalized transactions.
       * @throws {Error} AuthError without a usable session; DataError if pagination
       *   does not end within the page cap.
       */
      getTransactions: function (accountId, range) {
        requireSession();
        const maxPages = ns.Config.TRANSACTIONS.MAX_PAGES;
        let raw = [];
        let key;
        let pages = 0;
        do {
          const page = deps.client.getTransactions(accountId, {
            date_from: range.dateFrom, date_to: range.dateTo, continuation_key: key
          });
          raw = raw.concat(page.transactions || []);
          key = page.continuation_key || undefined;
          pages++;
        } while (key && pages < maxPages);
        if (key) throw ns.Errors.create('DataError', 'Transactions did not end within ' + maxPages + ' pages');
        return ns.EbNormalize.normalizeTransactions(raw, accountId).transactions;
      }
    });
  }

  ns.EnableBankingProvider = { create: create };
})(globalThis.BankSync);
