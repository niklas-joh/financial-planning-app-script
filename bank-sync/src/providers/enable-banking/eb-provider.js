/**
 * @fileoverview EnableBankingProvider: implements the BankProvider contract
 * for one bank over Enable Banking. Transactions arrive in Phase 2.
 */
(function (ns) {
  /**
   * @typedef {Object} EbProviderDeps
   * @property {import('../../core/config.js').BankConfig} bank - Bank this instance serves.
   * @property {Object} client - EbClient.
   * @property {Object} store - EbSessionStore.
   * @property {Object} auth - EbAuth flow.
   * @property {function(): number} nowMs - Clock.
   */

  /**
   * Creates a provider for one bank.
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
      const session = deps.store.loadSession(deps.bank.key);
      if (!session) throw ns.Errors.auth('Not authorized yet; run "Authorize" first');
      if (Date.parse(session.accessValidUntil) <= deps.nowMs()) {
        throw ns.Errors.auth('Consent has expired; re-authorize');
      }
      return session;
    }

    return ns.BankProvider.assertBankProvider({
      authorize: function () {
        return deps.auth.begin(deps.bank);
      },

      listAccounts: function () {
        const session = requireSession();
        const state = deps.client.getSession(session.sessionId);
        if (state.status && state.status !== 'AUTHORIZED') {
          throw ns.Errors.auth('Session status is ' + state.status + '; re-authorize');
        }
        return session.accounts;
      },

      getBalances: function (accountId) {
        requireSession();
        const fetchedAt = new Date(deps.nowMs()).toISOString();
        return (deps.client.getBalances(accountId).balances || []).map(function (raw) {
          return ns.EbNormalize.normalizeBalance(raw, accountId, fetchedAt);
        });
      },

      getTransactions: function () {
        throw ns.Errors.notImplemented('getTransactions');
      }
    });
  }

  ns.EnableBankingProvider = { create: create };
})(globalThis.BankSync);
