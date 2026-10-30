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
   * @property {{aspspName: string, country: string}} bank - Bank as listed by Enable Banking `GET /aspsps`
   *   (for Trading 212: name 'Trading 212', no country).
   * @property {string} [provider] - Provider id serving the connection; missing means 'enablebanking'
   *   (stored Enable Banking connections predate the field).
   * @property {string} [propertySuffix] - Trading 212 only: the holder part of the Script Property names
   *   (e.g. 'NIKLAS' in `T212_API_KEY_NIKLAS`). Never the secret itself.
   */

  /**
   * @typedef {Object} BankSyncConfig
   * @property {string} REPORTING_CURRENCY - ISO 4217 code all amounts are converted to.
   * @property {string} TIMEZONE - IANA timezone used for date handling.
   * @property {boolean} MASK_ACCOUNT_NUMBERS - When true, IBANs and other account, card and reference numbers
   *   are cut to their last four characters ('****1234') before anything is stored: the `ibanMasked` column,
   *   `counterpartyAccount`, `reference`, CSV texts and the `api.*` columns. When false (the default) they are
   *   stored in full. Every such call goes through `Redact.maskAccount`. Log and error text is always
   *   scrubbed by `Redact.text`, whatever this flag says. After turning it off, *Re-process stored rows*
   *   rewrites rows synced while it was on.
   * @property {Object} ENABLE_BANKING - API base URL, JWT claims/TTL, consent defaults.
   * @property {Object} LEGACY_CONNECTION - The single login that existed before connections were
   *   stored; used once to migrate its session. Connections are otherwise added with the
   *   Connect menu and kept in Script Properties.
   * @property {Object<string,string>} PROPERTY_KEYS - Script Properties key names.
   * @property {Object<string,string>} TABS - Sheet tab names.
   * @property {Object} ROUTING - Which of the three transaction tabs a row is stored on (see
   *   `core/txn-routing.js`): the providers whose rows are investment-account movements, and per provider
   *   the movement types (`description`) that stay unclassified until a report labels them (Trading 212
   *   card payments and cashback come through the API as plain WITHDRAW and DEPOSIT rows); and the CSV
   *   sources (format ids, the prefix of their account ids) whose rows are investment-account movements.
   * @property {Object<string,number>} TRANSACTIONS - Sync tuning: first-run lookback and
   *   re-fetch overlap in days, fetch window size in days, page cap per fetch, and the
   *   run time budget in ms (kept under the 6-minute Apps Script limit).
   * @property {Object} FX - Exchange rates: Frankfurter v2 base URL and rate provider (ECB), how many
   *   days a rate may be older than the date it converts (weekends and holidays), how long CacheService
   *   keeps the rates and the "already checked" markers, and the largest cache value in characters.
   * @property {Object} AUTOMATION - Daily sync: the global functions the daily and the follow-up
   *   triggers call, the hour of the daily run (script time zone), how long after an unfinished run
   *   the follow-up starts (ms), and how many follow-ups one daily run may chain; the monthly Trading 212
   *   report: the global functions that request it and fetch it, the day of the month and hour of the request
   *   (script time zone), how long after a request (or an unfinished check) the fetch runs (ms), and how many
   *   checks a requested report gets before the wait stops.
   * @property {Object} TRADING212 - Trading 212 public API: provider id, display name, origin and API path
   *   (live), page size (the API's maximum), pages per history feed the one-time backfill reads per run,
   *   the longest wait for a rate-limit reset before a call gives up; the report export: the start of the
   *   history when no Trading 212 row is stored yet, the most days one report may cover (Trading 212 refuses
   *   longer periods; the history is fetched one such period after the other), the days each report reaches
   *   back before the last one's end, the
   *   time after an export run starts during which a rate-limited call may still wait for its reset (ms; it
   *   bounds those waits, not the run), and how many redirects a report download follows.
   * @property {Object} CSV - CSV import: provider id, display name of its connections (one Drive folder per
   *   holder), the largest file read, and how many skipped files (and characters each) the last-read record
   *   names, so it stays far below the 9 KB Script Property limit.
   * @property {{MAX_LISTED: number}} REPROCESS - *Re-process stored rows*: how many rows still holding a masked value
   *   its summary names (the rest are counted).
   * @property {{WARN_DAYS: Array<number>}} CONSENT - Days before a consent lapses on which a warning
   *   email is sent (each once per consent); one more is sent once it has lapsed.
   */

  /** @type {BankSyncConfig} */
  ns.Config = Object.freeze({
    REPORTING_CURRENCY: 'EUR',
    TIMEZONE: 'Europe/Brussels',
    MASK_ACCOUNT_NUMBERS: false,
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
      EB_API_PREFIX: 'EB_API_',
      T212_API_KEY_PREFIX: 'T212_API_KEY_',
      T212_API_SECRET_PREFIX: 'T212_API_SECRET_',
      T212_HISTORY_PREFIX: 'T212_HISTORY_',
      T212_EXPORT_PREFIX: 'T212_EXPORT_',
      T212_EXPORT_FOLDER_PREFIX: 'CSV_FOLDER_T212_',
      CSV_FOLDER_PREFIX: 'CSV_FOLDER_',
      CSV_IMPORTS_PREFIX: 'CSV_IMPORTS_',
      LAST_RUN: 'BS_LAST_RUN',
      CONSENT_WARNED_PREFIX: 'BS_CONSENT_WARNED_'
    }),
    TABS: Object.freeze({
      RAW_ACCOUNTS: 'raw_accounts',
      RAW_BALANCES: 'raw_balances',
      RAW_TRANSACTIONS: 'raw_transactions',
      TRANSACTIONS: 'transactions',
      INVESTMENT_TRANSACTIONS: 'investment_transactions',
      UNCLASSIFIED_TRANSACTIONS: 'unclassified_transactions',
      RATES: 'rates',
      HOLDINGS: 'holdings',
      DAILY_VALUES: 'daily_values'
    }),
    ROUTING: Object.freeze({
      INVESTMENT_PROVIDERS: Object.freeze({
        trading212: Object.freeze({ unclearTypes: Object.freeze(['WITHDRAW', 'DEPOSIT']) })
      }),
      INVESTMENT_CSV_SOURCES: Object.freeze(['avanza'])
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
    }),
    FX: Object.freeze({
      BASE_URL: 'https://api.frankfurter.dev/v2',
      RATE_PROVIDER: 'ecb',
      SOURCE_LABEL: 'ECB',
      MAX_FALLBACK_DAYS: 7,
      CACHE_SECONDS: 21600,
      CACHE_MAX_CHARS: 90000
    }),
    AUTOMATION: Object.freeze({
      DAILY_HANDLER: 'triggerDailySync',
      CONTINUE_HANDLER: 'triggerSyncContinue',
      DAILY_HOUR: 6,
      CONTINUE_AFTER_MS: 300000,
      MAX_CONTINUATIONS: 3,
      REPORT_REQUEST_HANDLER: 'triggerT212ReportRequest',
      REPORT_FETCH_HANDLER: 'triggerT212ReportFetch',
      REPORT_DAY: 1,
      REPORT_HOUR: 5,
      REPORT_FETCH_AFTER_MS: 300000,
      REPORT_MAX_CHECKS: 6
    }),
    TRADING212: Object.freeze({
      PROVIDER_ID: 'trading212',
      BANK_NAME: 'Trading 212',
      ORIGIN: 'https://live.trading212.com',
      API_PATH: '/api/v0',
      PAGE_LIMIT: 50,
      BACKFILL_PAGES_PER_RUN: 40,
      MAX_RETRY_WAIT_MS: 65000,
      EXPORT_FIRST_FROM: '2016-01-01',
      EXPORT_MAX_DAYS: 365,
      EXPORT_OVERLAP_DAYS: 7,
      EXPORT_BUDGET_MS: 240000,
      MAX_REDIRECTS: 3
    }),
    CSV: Object.freeze({
      PROVIDER_ID: 'csv',
      BANK_NAME: 'CSV import',
      MAX_FILE_BYTES: 5000000,
      MAX_SKIPPED_LISTED: 10,
      MAX_SKIPPED_CHARS: 200
    }),
    REPROCESS: Object.freeze({
      MAX_LISTED: 25
    }),
    CONSENT: Object.freeze({
      WARN_DAYS: Object.freeze([30, 14, 7, 1])
    })
  });
})(globalThis.BankSync);
