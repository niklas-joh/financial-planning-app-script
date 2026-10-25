/**
 * @fileoverview Static, non-secret configuration. Secrets (application ID,
 * private key) are never stored here; they are read from Script Properties
 * under the keys named in `PROPERTY_KEYS`.
 */
(function (ns) {
  /**
   * @typedef {Object} ConnectionConfig
   * @property {string} key - Stable key, bank plus holder (e.g. 'revolut-niklas'); one login and consent.
   * @property {string} holder - Short label of the person the login belongs to.
   * @property {{aspspName: string, country: string}} bank - Bank as listed by Enable Banking `GET /aspsps`.
   */

  /**
   * @typedef {Object} BankSyncConfig
   * @property {string} REPORTING_CURRENCY - ISO 4217 code all amounts are converted to.
   * @property {string} TIMEZONE - IANA timezone used for date handling.
   * @property {Object} ENABLE_BANKING - API base URL, JWT claims/TTL, consent defaults.
   * @property {Object} LEGACY_CONNECTION - The single login that existed before connections were
   *   stored; used once to migrate its session. Connections are otherwise added with the
   *   Connect menu and kept in Script Properties.
   * @property {Object<string,string>} PROPERTY_KEYS - Script Properties key names.
   * @property {Object<string,string>} TABS - Sheet tab names.
   * @property {Object<string,number>} TRANSACTIONS - Sync tuning: first-run lookback and
   *   re-fetch overlap in days, fetch window size in days, page cap per fetch, and the
   *   run time budget in ms (kept under the 6-minute Apps Script limit).
   */

  /** @type {BankSyncConfig} */
  ns.Config = Object.freeze({
    REPORTING_CURRENCY: 'EUR',
    TIMEZONE: 'Europe/Luxembourg',
    ENABLE_BANKING: Object.freeze({
      PROVIDER_ID: 'enablebanking',
      BASE_URL: 'https://api.enablebanking.com',
      JWT_ISSUER: 'enablebanking.com',
      JWT_AUDIENCE: 'api.enablebanking.com',
      JWT_TTL_SECONDS: 3600,
      DEFAULT_CONSENT_DAYS: 90,
      PENDING_AUTH_TTL_SECONDS: 900
    }),
    LEGACY_CONNECTION: Object.freeze({ sessionKey: 'revolut', bankName: 'Revolut', country: 'LU', holder: 'niklas' }),
    PROPERTY_KEYS: Object.freeze({
      EB_APP_ID: 'EB_APP_ID',
      EB_PRIVATE_KEY: 'EB_PRIVATE_KEY',
      EB_REDIRECT_URL: 'EB_REDIRECT_URL',
      EB_SESSIONS: 'EB_SESSIONS',
      EB_PENDING: 'EB_PENDING',
      EB_CONNECTIONS: 'EB_CONNECTIONS',
      EB_MIGRATED: 'EB_MIGRATED',
      EB_SESSION_PREFIX: 'EB_SESSION_',
      EB_API_PREFIX: 'EB_API_'
    }),
    TABS: Object.freeze({
      RAW_ACCOUNTS: 'raw_accounts',
      RAW_BALANCES: 'raw_balances',
      RAW_TRANSACTIONS: 'raw_transactions',
      TRANSACTIONS: 'transactions'
    }),
    TRANSACTIONS: Object.freeze({
      INITIAL_LOOKBACK_DAYS: 90,
      OVERLAP_DAYS: 7,
      WINDOW_DAYS: 31,
      MAX_PAGES: 50,
      LOCK_WAIT_MS: 30000,
      MIN_CONNECTION_BUDGET_MS: 15000,
      RAW_RETENTION_DAYS: 365,
      DETAILS_MAX_AGE_DAYS: 7,
      TIME_BUDGET_MS: 270000
    })
  });
})(globalThis.BankSync);
