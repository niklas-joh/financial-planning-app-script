/**
 * @fileoverview The user-authorization flow: start (POST /auth) and complete
 * (POST /sessions). Consent expiry is requested here and persisted.
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
   * @property {function(): string} getRedirectUrl - Registered redirect URL.
   * @property {function(): number} nowMs - Clock.
   * @property {function(): string} uuid - Random state generator.
   */

  /**
   * Creates the auth flow.
   * @param {AuthDeps} deps - Injected services.
   * @returns {{begin: function(Object): {url: string, state: string}, complete: function(string, string): Object}} Flow.
   */
  function create(deps) {
    return {
      /**
       * Starts authorization for a bank.
       * @param {import('../../core/config.js').BankConfig} bank - Bank to connect.
       * @returns {{url: string, state: string}} Login URL the user must open.
       * @throws {Error} ConfigError if the bank or redirect URL is unavailable.
       */
      begin: function (bank) {
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
        deps.store.savePending(state, { bankKey: bank.key, validUntil: validUntil });
        return { url: response.url, state: state };
      },

      /**
       * Finishes authorization after the bank redirects back.
       * @param {string} code - `code` query parameter.
       * @param {string} state - `state` query parameter.
       * @returns {{bankKey: string, accountCount: number, validUntil: string}} Non-secret summary.
       * @throws {Error} AuthError for an unknown/expired state.
       */
      complete: function (code, state) {
        const pending = deps.store.consumePending(state);
        const session = deps.client.createSession(code);
        const validUntil = (session.access && session.access.valid_until) || pending.validUntil;
        const accounts = (session.accounts || []).map(ns.EbNormalize.normalizeAccount);
        deps.store.saveSession(pending.bankKey, {
          sessionId: session.session_id,
          accessValidUntil: validUntil,
          createdAt: new Date(deps.nowMs()).toISOString(),
          accounts: accounts
        });
        return { bankKey: pending.bankKey, accountCount: accounts.length, validUntil: validUntil };
      }
    };
  }

  ns.EbAuth = { create: create, findAspsp: findAspsp, consentExpiry: consentExpiry };
})(globalThis.BankSync);
