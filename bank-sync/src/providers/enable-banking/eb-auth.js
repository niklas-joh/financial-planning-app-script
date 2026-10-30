/**
 * @fileoverview The user-authorization flow: start (POST /auth) and complete
 * (POST /sessions), plus a user-triggered refresh of account details. Consent
 * expiry is requested here and persisted. Sessions are kept per connection
 * (one login at one bank), so several people can use the same bank.
 */
(function (ns) {
  const DAY_MS = 86400000;
  const SAFETY_MARGIN_MS = 5 * 60000;

  /**
   * Finds a bank in an ASPSP listing by exact (case-insensitive) name.
   * @param {Array<{name: string}>} aspsps - Listing from `GET /aspsps`.
   * @param {string} name - Wanted bank name.
   * @returns {Object} The matching ASPSP.
   * @throws {Error} ConfigError listing similar names if not found.
   */
  function findAspsp(aspsps, name) {
    const wanted = name.toLowerCase();
    const match = aspsps.find(function (a) { return a.name.toLowerCase() === wanted; });
    if (match) return match;
    const similar = aspsps
      .filter(function (a) { return a.name.toLowerCase().indexOf(wanted.slice(0, 4)) !== -1; })
      .map(function (a) { return a.name; });
    throw ns.Errors.config('Bank "' + name + '" not listed; similar: ' + (similar.join(', ') || 'none'));
  }

  /**
   * Computes the consent expiry, clamped to the bank's maximum.
   * @param {number} nowMs - Current time in epoch milliseconds.
   * @param {number} requestedDays - Desired consent length in days.
   * @param {?number} maxSeconds - Bank's `maximum_consent_validity`, if known.
   * @returns {string} ISO 8601 UTC timestamp, just inside the allowed window.
   */
  function consentExpiry(nowMs, requestedDays, maxSeconds) {
    const requestedMs = requestedDays * DAY_MS;
    const allowedMs = maxSeconds ? Math.min(requestedMs, maxSeconds * 1000) : requestedMs;
    return new Date(nowMs + allowedMs - SAFETY_MARGIN_MS).toISOString();
  }

  /**
   * @typedef {Object} AuthDeps
   * @property {Object} client - EbClient.
   * @property {Object} store - EbSessionStore.
   * @property {{save: function(Object): void}} connections - Connection store.
   * @property {function(): string} getRedirectUrl - Registered redirect URL.
   * @property {function(): number} nowMs - Clock.
   * @property {function(): string} uuid - Random state generator.
   */

  /**
   * Removes the bulky API fields before an account goes into the session.
   * @param {import('../bank-provider.js').Account} account - Account with `apiFields`.
   * @returns {import('../bank-provider.js').Account} Copy without `apiFields`.
   */
  function withoutApiFields(account) {
    const copy = Object.assign({}, account);
    delete copy.apiFields;
    return copy;
  }

  /**
   * Creates the auth flow.
   * @param {AuthDeps} deps - Injected services.
   * @returns {{begin: function(Object): {url: string, state: string}, complete: function(string, string): Object,
   *   refreshStale: function(Object, function(): boolean=, boolean=): Array<Object>}} Flow.
   */
  function create(deps) {
    /**
     * Normalizes full AccountResources, keeps their API fields and returns the lean accounts.
     * @param {Array<Object>} resources - AccountResources.
     * @param {boolean} detailsRead - True if they came from the details endpoint; otherwise
     *   `detailsAt` stays empty so the first sync reads the details once.
     * @returns {Array<import('../bank-provider.js').Account>} Accounts for the session.
     */
    function keepAccounts(resources, detailsRead) {
      const at = detailsRead ? new Date(deps.nowMs()).toISOString() : '';
      return resources.map(function (raw) {
        const account = ns.EbNormalize.normalizeAccount(raw, at);
        deps.store.saveApiFields(account.accountId, account.apiFields);
        return withoutApiFields(account);
      });
    }

    /**
     * Whether an account's details are missing or older than the allowed age.
     * @param {import('../bank-provider.js').Account} account - Stored account.
     * @returns {boolean} True if they should be read again.
     */
    function isStale(account) {
      const maxAgeMs = ns.Config.TRANSACTIONS.DETAILS_MAX_AGE_DAYS * DAY_MS;
      return !account.detailsAt || deps.nowMs() - Date.parse(account.detailsAt) > maxAgeMs;
    }

    return {
      /**
       * Starts authorization for a connection.
       * @param {import('../../core/config.js').ConnectionConfig} connection - Login to connect.
       * @returns {{url: string, state: string}} Login URL the user must open.
       * @throws {Error} ConfigError if the bank or redirect URL is unavailable; HttpError from
       *   Enable Banking (see `Connections.explainNotEnabled`).
       */
      begin: function (connection) {
        const bank = connection.bank;
        const redirectUrl = deps.getRedirectUrl();
        if (!redirectUrl) throw ns.Errors.config('Script Property EB_REDIRECT_URL is not set');
        const aspsp = findAspsp(deps.client.listAspsps(bank.country).aspsps, bank.aspspName);
        const validUntil = consentExpiry(
          deps.nowMs(), ns.Config.ENABLE_BANKING.DEFAULT_CONSENT_DAYS, aspsp.maximum_consent_validity);
        const state = deps.uuid();
        const response = deps.client.startAuth({
          access: { valid_until: validUntil },
          aspsp: { name: aspsp.name, country: aspsp.country },
          state: state,
          redirect_url: redirectUrl,
          psu_type: 'personal'
        });
        deps.store.savePending(state, { connection: connection, validUntil: validUntil });
        return { url: response.url, state: state };
      },

      /**
       * Finishes authorization after the bank redirects back. The connection is stored only
       * once the bank has granted at least one account.
       * @param {string} code - `code` query parameter.
       * @param {string} state - `state` query parameter.
       * @returns {{connectionKey: string, accountCount: number, validUntil: string}} Non-secret summary.
       * @throws {Error} AuthError for an unknown/expired state, or if no account is enabled
       *   for the bank (`apiCode` NO_ACCOUNTS_ADDED).
       */
      complete: function (code, state) {
        const pending = deps.store.consumePending(state);
        const connection = pending.connection;
        if (!connection) throw ns.Errors.auth('Authorization started before an update; start again');
        const session = deps.client.createSession(code);
        const resources = session.accounts || [];
        if (!resources.length) {
          throw ns.Errors.create('AuthError', 'The bank granted no accounts', { apiCode: 'NO_ACCOUNTS_ADDED' });
        }
        const validUntil = (session.access && session.access.valid_until) || pending.validUntil;
        const accounts = keepAccounts(resources, false);
        deps.connections.save(connection);
        deps.store.saveSession(connection.key, {
          sessionId: session.session_id,
          accessValidUntil: validUntil,
          createdAt: new Date(deps.nowMs()).toISOString(),
          accounts: accounts
        });
        return { connectionKey: connection.key, accountCount: accounts.length, validUntil: validUntil };
      },

      /**
       * Reads the full details of accounts whose details are missing or older than
       * `DETAILS_MAX_AGE_DAYS` (one bank call each), so names set by the user or bank and
       * every API field stay current without any manual step. A transient failure keeps the old
       * account and is retried on the next sync; a 401/403 (consent problem) is re-thrown.
       * @param {import('../../core/config.js').ConnectionConfig} connection - Connected login.
       * @param {function(): boolean} [outOfTime] - True once the shared time budget is spent; the
       *   remaining stale accounts are then left for the next sync.
       * @param {boolean} [force] - Read every account's details now, however fresh (re-process).
       * @returns {Array<import('../bank-provider.js').Account>} The session's accounts, refreshed where due.
       * @throws {Error} AuthError if the connection has no session; HttpError 401/403 from the bank.
       */
      refreshStale: function (connection, outOfTime, force) {
        const session = deps.store.loadSession(connection.key);
        if (!session) throw ns.Errors.auth('Not authorized yet; connect first');
        let changed = false;
        const accounts = session.accounts.map(function (account) {
          if ((!force && !isStale(account)) || (outOfTime && outOfTime())) return account;
          try {
            const fresh = keepAccounts([deps.client.getAccountDetails(account.accountId)], true)[0];
            changed = true;
            return fresh;
          } catch (e) {
            if (e.status === 401 || e.status === 403) throw e;
            return account;
          }
        });
        if (changed) deps.store.saveSession(connection.key, Object.assign({}, session, { accounts: accounts }));
        return accounts;
      }
    };
  }

  ns.EbAuth = { create: create, findAspsp: findAspsp, consentExpiry: consentExpiry };
})(globalThis.BankSync);
