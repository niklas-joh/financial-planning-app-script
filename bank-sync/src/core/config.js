/**
 * @fileoverview Static, non-secret configuration. Secrets (application ID,
 * private key) are never stored here; they are read from Script Properties
 * under the keys named in `PROPERTY_KEYS`.
 */
(function (ns) {
  /**
   * @typedef {Object} BankConfig
   * @property {string} key - Stable internal key for the bank.
   * @property {string} aspspName - Name as listed by Enable Banking `GET /aspsps`.
   * @property {string} country - ISO 3166 country the bank is listed under.
   */

  /**
   * @typedef {Object} ConnectionConfig
   * @property {string} key - Stable key, bank plus holder (e.g. 'revolut-niklas'); one login and consent.
   * @property {string} holder - Short label of the person the login belongs to.
   * @property {BankConfig} bank - Bank the login is at.
   * @property {string} [legacySessionKey] - Key an earlier version stored this session under.
   */

  /**
   * @typedef {Object} BankSyncConfig
   * @property {string} REPORTING_CURRENCY - ISO 4217 code all amounts are converted to.
   * @property {string} TIMEZONE - IANA timezone used for date handling.
   * @property {Object} ENABLE_BANKING - API base URL, JWT claims/TTL, consent defaults.
   * @property {Object<string,BankConfig>} BANKS - Banks reachable through Enable Banking.
   * @property {Object<string,ConnectionConfig>} CONNECTIONS - Logins to sync. Adding a bank or
   *   another person at a bank is a new entry here, not a code change.
   * @property {Object<string,string>} PROPERTY_KEYS - Script Properties key names.
   * @property {Object<string,string>} TABS - Sheet tab names.
   * @property {Object<string,number>} TRANSACTIONS - Sync tuning: first-run lookback and
   *   re-fetch overlap in days, fetch window size in days, page cap per fetch, and the
   *   run time budget in ms (kept under the 6-minute Apps Script limit).
   */

  const BANKS = Object.freeze({
    REVOLUT: Object.freeze({ key: 'revolut', aspspName: 'Revolut', country: 'LU' })
  });

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
    BANKS: BANKS,
    CONNECTIONS: Object.freeze({
      REVOLUT_NIKLAS: Object.freeze({
        key: 'revolut-niklas', holder: 'niklas', bank: BANKS.REVOLUT, legacySessionKey: 'revolut'
      })
    }),
    PROPERTY_KEYS: Object.freeze({
      EB_APP_ID: 'EB_APP_ID',
      EB_PRIVATE_KEY: 'EB_PRIVATE_KEY',
      EB_REDIRECT_URL: 'EB_REDIRECT_URL',
      EB_SESSIONS: 'EB_SESSIONS',
      EB_PENDING: 'EB_PENDING'
    }),
    TABS: Object.freeze({
      RAW_ACCOUNTS: 'raw_accounts',
      RAW_BALANCES: 'raw_balances',
      TRANSACTIONS: 'transactions'
    }),
    TRANSACTIONS: Object.freeze({
      INITIAL_LOOKBACK_DAYS: 90,
      OVERLAP_DAYS: 7,
      WINDOW_DAYS: 31,
      MAX_PAGES: 50,
      LOCK_WAIT_MS: 30000,
      MIN_CONNECTION_BUDGET_MS: 15000,
      TIME_BUDGET_MS: 270000
    })
  });
})(globalThis.BankSync);
