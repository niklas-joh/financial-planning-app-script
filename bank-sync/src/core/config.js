/**
 * @fileoverview Static, non-secret configuration. Secrets (application ID,
 * private key) are never stored here; they are read from Script Properties
 * under the keys named in `PROPERTY_KEYS`.
 */
(function (root) {
  var ns = root.BankSync = root.BankSync || {};

  /**
   * @typedef {Object} EnableBankingConfig
   * @property {string} BASE_URL - REST API base URL.
   * @property {string} JWT_ISSUER - Required `iss` claim.
   * @property {string} JWT_AUDIENCE - Required `aud` claim.
   * @property {number} JWT_TTL_SECONDS - Lifetime of each signed JWT.
   */

  /**
   * @typedef {Object} BankSyncConfig
   * @property {string} REPORTING_CURRENCY - ISO 4217 code all amounts are converted to.
   * @property {string} TIMEZONE - IANA timezone used for date handling.
   * @property {EnableBankingConfig} ENABLE_BANKING - Enable Banking API settings.
   * @property {Object<string,string>} PROPERTY_KEYS - Script Properties key names.
   * @property {Object<string,string>} TABS - Sheet tab names.
   */

  /** @type {BankSyncConfig} */
  var Config = Object.freeze({
    REPORTING_CURRENCY: 'EUR',
    TIMEZONE: 'Europe/Luxembourg',
    ENABLE_BANKING: Object.freeze({
      BASE_URL: 'https://api.enablebanking.com',
      JWT_ISSUER: 'enablebanking.com',
      JWT_AUDIENCE: 'api.enablebanking.com',
      JWT_TTL_SECONDS: 3600
    }),
    PROPERTY_KEYS: Object.freeze({
      EB_APP_ID: 'EB_APP_ID',
      EB_PRIVATE_KEY: 'EB_PRIVATE_KEY',
      EB_REDIRECT_URL: 'EB_REDIRECT_URL'
    }),
    TABS: Object.freeze({
      RAW_ACCOUNTS: 'raw_accounts',
      RAW_BALANCES: 'raw_balances'
    })
  });

  ns.Config = Config;
})(typeof globalThis !== 'undefined' ? globalThis : this);

if (typeof module !== 'undefined' && module.exports) {
  module.exports = globalThis.BankSync.Config;
}
