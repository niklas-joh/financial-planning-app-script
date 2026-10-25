/**
 * @fileoverview Persists pending authorizations and authorized sessions in
 * Script Properties (private to the script owner). `props` is injected.
 */
(function (ns) {
  const KEYS = ns.Config.PROPERTY_KEYS;
  const MAX_PROPERTY_CHARS = 8000;
  const PENDING_TTL_MS = ns.Config.ENABLE_BANKING.PENDING_AUTH_TTL_SECONDS * 1000;

  /**
   * @typedef {Object} StoredSession
   * @property {string} sessionId - Enable Banking session id.
   * @property {string} accessValidUntil - ISO date-time when consent lapses.
   * @property {string} createdAt - ISO date-time the session was created.
   * @property {Array<import('../bank-provider.js').Account>} accounts - Accounts granted (IBANs masked; the
   *   full API fields live in `loadApiFields`, not here).
   */

  /**
   * Creates a store.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): *, deleteProperty: function(string): *}} props - Script Properties.
   * @param {function(): number} nowMs - Clock in epoch milliseconds.
   * @returns {Object} Store with savePending, consumePending, saveSession, loadSession, migrateKey, saveApiFields, loadApiFields.
   */
  function create(props, nowMs) {
    function readMap(key) {
      const raw = props.getProperty(key);
      return raw ? JSON.parse(raw) : {};
    }

    function writeMap(key, map) {
      props.setProperty(key, JSON.stringify(map));
    }

    return {
      /**
       * Remembers a started authorization, pruning expired ones.
       * @param {string} state - Random state sent to the bank.
       * @param {Object} data - Context needed to finish (bank key, requested expiry).
       */
      savePending: function (state, data) {
        const now = nowMs();
        const map = readMap(KEYS.EB_PENDING);
        Object.keys(map).forEach(function (s) {
          if (map[s].expiresAt <= now) delete map[s];
        });
        map[state] = Object.assign({}, data, { expiresAt: now + PENDING_TTL_MS });
        writeMap(KEYS.EB_PENDING, map);
      },

      /**
       * Returns and removes a pending authorization (single use).
       * @param {string} state - State echoed back by the bank.
       * @returns {Object} The saved context.
       * @throws {Error} AuthError if unknown or expired.
       */
      consumePending: function (state) {
        const map = readMap(KEYS.EB_PENDING);
        const entry = map[state];
        if (!entry || entry.expiresAt <= nowMs()) {
          throw ns.Errors.auth('Unknown or expired authorization state; start again');
        }
        delete map[state];
        writeMap(KEYS.EB_PENDING, map);
        return entry;
      },

      /**
       * Stores a session in its own property (Script Properties values are limited to 9 KB,
       * so sessions of several connections must not share one).
       * @param {string} connectionKey - Connection key.
       * @param {StoredSession} session - Session to persist.
       */
      saveSession: function (connectionKey, session) {
        props.setProperty(KEYS.EB_SESSION_PREFIX + connectionKey, JSON.stringify(session));
      },

      /**
       * @param {string} connectionKey - Connection key.
       * @returns {?StoredSession} The session, or null if never authorized. Falls back to the
       *   single-map storage used before sessions had their own property.
       */
      loadSession: function (connectionKey) {
        const own = props.getProperty(KEYS.EB_SESSION_PREFIX + connectionKey);
        return own ? JSON.parse(own) : readMap(KEYS.EB_SESSIONS)[connectionKey] || null;
      },

      /**
       * Moves a session stored under an old key to its new key. Idempotent; never
       * overwrites a session that already exists under the new key.
       * @param {string} fromKey - Old key.
       * @param {string} toKey - New key.
       * @returns {boolean} True if a session was moved.
       */
      migrateKey: function (fromKey, toKey) {
        const session = this.loadSession(fromKey);
        if (!session || this.loadSession(toKey)) return false;
        this.saveSession(toKey, session);
        const legacy = readMap(KEYS.EB_SESSIONS);
        if (legacy[fromKey]) {
          delete legacy[fromKey];
          writeMap(KEYS.EB_SESSIONS, legacy);
        }
        props.deleteProperty(KEYS.EB_SESSION_PREFIX + fromKey);
        return true;
      },

      /**
       * Keeps an account's full flattened API fields (shown on the raw tabs), one property
       * per account so they never compete with the session for space.
       * @param {string} accountId - Account uid.
       * @param {Object<string,*>} fields - Flat, IBAN-masked fields.
       * @returns {boolean} False if the fields are too large for one property and were not stored
       *   (the account then has no `api.*` columns; it never blocks a connect or a sync).
       */
      saveApiFields: function (accountId, fields) {
        const json = JSON.stringify(fields);
        if (json.length > MAX_PROPERTY_CHARS) return false;
        props.setProperty(KEYS.EB_API_PREFIX + accountId, json);
        return true;
      },

      /**
       * @param {string} accountId - Account uid.
       * @returns {Object<string,*>} The stored fields, or {} if none.
       */
      loadApiFields: function (accountId) {
        const raw = props.getProperty(KEYS.EB_API_PREFIX + accountId);
        return raw ? JSON.parse(raw) : {};
      }
    };
  }

  ns.EbSessionStore = { create: create };
})(globalThis.BankSync);
