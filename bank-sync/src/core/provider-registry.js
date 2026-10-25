/**
 * @fileoverview The providers Bank Sync knows, behind one small interface each: which
 * connections it has, whether a connection can sync, how to build its provider, what Show
 * status says about it, and its consents (for the expiry warnings). Sync all, Show status and
 * the consent warnings loop over the entries, so adding a provider means adding an entry here
 * and nothing in the menu, Sync all or another provider. Apps Script services arrive as
 * arguments (`wired` from `Container.services()`, `env` from the container), so this stays
 * Node-testable.
 */
(function (ns) {
  /**
   * @typedef {Object} ProviderEnv
   * @property {function(Object=): {request: function(Object): *}} http - HTTP client factory; options are
   *   extra `Http.create` deps (e.g. `retryDelayMs`).
   * @property {function(string): string} base64 - Base64 encoder (UTF-8).
   * @property {function(string): string} toLocalDate - ISO date-time to YYYY-MM-DD in the Bank Sync time zone.
   * @property {function(): number} [msLeft] - Time left in the current sync run.
   */

  /**
   * @typedef {Object} ProviderEntry
   * @property {string} id - Provider id (`ConnectionConfig.provider`).
   * @property {function(Object): Array<import('./config.js').ConnectionConfig>} connections - Its connections.
   * @property {function(import('./config.js').ConnectionConfig, Object): boolean} isConnected - Ready to sync.
   * @property {function(import('./config.js').ConnectionConfig, Object, ProviderEnv=): import('../providers/bank-provider.js').BankProvider} create
   * @property {function(Object): Array<string>} statusLines - Show status lines (names and set/MISSING only).
   * @property {function(Object): Array<{connection: Object, validUntil: string}>} consents - Consents that expire.
   */

  /** @type {ProviderEntry} */
  const enableBanking = {
    id: ns.Config.ENABLE_BANKING.PROVIDER_ID,
    connections: function (wired) { return wired.connections.all(); },
    isConnected: function (c, wired) { return Boolean(wired.store.loadSession(c.key)); },
    create: function (c, wired) {
      return ns.EnableBankingProvider.create({ connection: c, client: wired.client, store: wired.store, auth: wired.auth, nowMs: wired.nowMs });
    },
    consents: function (wired) {
      return wired.connections.all().map(function (c) {
        const session = wired.store.loadSession(c.key);
        return session ? { connection: c, validUntil: session.accessValidUntil } : null;
      }).filter(Boolean);
    },
    statusLines: function (wired) {
      const keys = ns.Config.PROPERTY_KEYS;
      const yesNo = function (key) { return key + ': ' + (wired.props.getProperty(key) ? 'set' : 'MISSING'); };
      const consentByKey = {};
      enableBanking.consents(wired).forEach(function (x) { consentByKey[x.connection.key] = x; });
      const connections = wired.connections.all().map(function (c) {
        const found = consentByKey[c.key];
        const now = wired.nowMs();
        const consent = found
          ? found.validUntil + ' (' + ns.ConsentWarnings.daysLeft(found.validUntil, now) + ' days left)' +
            ns.ConsentWarnings.statusNote(found.validUntil, now, ns.Config.CONSENT.WARN_DAYS)
          : 'not connected';
        return ns.Connections.label(c) + ': ' + consent;
      });
      return [yesNo(keys.EB_APP_ID), yesNo(keys.EB_PRIVATE_KEY), yesNo(keys.EB_REDIRECT_URL)]
        .concat(connections.length ? connections : ['No bank connected yet: use "Connect a bank..."']);
    }
  };

  /** @type {ProviderEntry} */
  const trading212 = {
    id: ns.Config.TRADING212.PROVIDER_ID,
    connections: function (wired) { return ns.T212Credentials.connections(wired.props); },
    isConnected: function (c, wired) { return ns.T212Credentials.isComplete(wired.props, c); },
    create: function (c, wired, env) {
      const http = env.http({ retryDelayMs: ns.T212RateLimit.hook(wired.nowMs, env.msLeft) });
      const client = ns.T212Client.create(http, ns.T212Credentials.credentials(wired.props, c), env.base64);
      return ns.Trading212Provider.create({ connection: c, client: client, nowMs: wired.nowMs, toLocalDate: env.toLocalDate });
    },
    consents: function () { return []; },
    statusLines: function (wired) { return ns.T212Credentials.statusLines(wired.props); }
  };

  const ENTRIES = [enableBanking, trading212];

  /**
   * @param {import('./config.js').ConnectionConfig} connection - Connection.
   * @returns {ProviderEntry} Its provider's entry (stored connections without `provider` are Enable Banking).
   * @throws {Error} ConfigError for an unknown provider id.
   */
  function entryFor(connection) {
    const id = connection.provider || enableBanking.id;
    const entry = ENTRIES.filter(function (e) { return e.id === id; })[0];
    if (!entry) throw ns.Errors.config('Unknown provider: ' + id);
    return entry;
  }

  /**
   * @param {string} method - Entry method returning a list.
   * @returns {function(Object): Array<*>} Concatenation of that list over all entries.
   */
  function across(method) {
    return function (wired) {
      return ENTRIES.reduce(function (all, entry) { return all.concat(entry[method](wired)); }, []);
    };
  }

  ns.ProviderRegistry = {
    entryFor: entryFor,
    /** Every connection of every provider, Enable Banking first. */
    connections: across('connections'),
    /** Show status lines of every provider. */
    statusLines: across('statusLines'),
    /** Every consent that expires (for the warnings). */
    consents: across('consents'),
    /**
     * @param {import('./config.js').ConnectionConfig} c - Connection.
     * @param {Object} wired - Services.
     * @returns {boolean} Whether it can sync.
     */
    isConnected: function (c, wired) { return entryFor(c).isConnected(c, wired); },
    /**
     * @param {import('./config.js').ConnectionConfig} c - Connection.
     * @param {Object} wired - Services.
     * @param {ProviderEnv} [env] - Apps Script helpers (needed by Trading 212).
     * @returns {import('../providers/bank-provider.js').BankProvider} Provider.
     */
    create: function (c, wired, env) { return entryFor(c).create(c, wired, env); }
  };
})(globalThis.BankSync);
