/**
 * @fileoverview Persists pending authorizations and authorized sessions in
 * Script Properties (private to the script owner). `props` is injected.
 */
(function (ns) {
  const KEYS = ns.Config.PROPERTY_KEYS;
  const PENDING_TTL_MS = ns.Config.ENABLE_BANKING.PENDING_AUTH_TTL_SECONDS * 1000;

  /**
   * @typedef {Object} StoredSession
   * @property {string} sessionId - Enable Banking session id.
   * @property {string} accessValidUntil - ISO date-time when consent lapses.
   * @property {string} createdAt - ISO date-time the session was created.
   * @property {Array<import('../bank-provider.js').Account>} accounts - Accounts granted (IBANs masked).
   */

  /**
   * Creates a store.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): *}} props - Script Properties.
   * @param {function(): number} nowMs - Clock in epoch milliseconds.
   * @returns {Object} Store with savePending, consumePending, saveSession, loadSession, migrateKey.
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
       * @param {string} bankKey - Connection key from Config.CONNECTIONS.
       * @param {StoredSession} session - Session to persist.
       */
      saveSession: function (bankKey, session) {
        const map = readMap(KEYS.EB_SESSIONS);
        map[bankKey] = session;
        writeMap(KEYS.EB_SESSIONS, map);
      },

      /**
       * @param {string} bankKey - Connection key from Config.CONNECTIONS.
       * @returns {?StoredSession} The session, or null if never authorized.
       */
      loadSession: function (bankKey) {
        return readMap(KEYS.EB_SESSIONS)[bankKey] || null;
      },

      /**
       * Moves a session stored under an old key to its new key. Idempotent; never
       * overwrites a session that already exists under the new key.
       * @param {string} fromKey - Old key.
       * @param {string} toKey - New key.
       * @returns {boolean} True if a session was moved.
       */
      migrateKey: function (fromKey, toKey) {
        const map = readMap(KEYS.EB_SESSIONS);
        if (!map[fromKey] || map[toKey]) return false;
        map[toKey] = map[fromKey];
        delete map[fromKey];
        writeMap(KEYS.EB_SESSIONS, map);
        return true;
      }
    };
  }

  ns.EbSessionStore = { create: create };
})(globalThis.BankSync);
