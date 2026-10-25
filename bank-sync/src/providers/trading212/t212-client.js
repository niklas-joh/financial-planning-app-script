/**
 * @fileoverview Thin Trading 212 public API client: one method per endpoint, no
 * business logic. Read-only endpoints only. Auth is HTTP Basic with the API key and
 * secret; the header is built here and never leaves this module.
 */
(function (ns) {
  const T212 = ns.Config.TRADING212;

  /**
   * URL for an API path. `nextPagePath` from a paged response already starts with the API path
   * (`/api/v0/...`) and is used as is; anything else is relative to it. Only paths are accepted,
   * so the key is never sent to another host.
   * @param {string} path - '/equity/...' or '/api/v0/equity/...'.
   * @returns {string} Absolute URL on the live origin.
   * @throws {Error} DataError if the path is not a path.
   */
  function urlFor(path) {
    const text = String(path || '');
    if (text.charAt(0) !== '/' || text.charAt(1) === '/') throw ns.Errors.create('DataError', 'Unexpected Trading 212 page path');
    return T212.ORIGIN + (text.indexOf(T212.API_PATH + '/') === 0 ? text : T212.API_PATH + text);
  }

  /**
   * Creates a client.
   * @param {{request: function(Object): *}} http - JSON HTTP client (with the rate-limit retry hook).
   * @param {{key: string, secret: string}} credentials - API key and secret.
   * @param {function(string): string} base64 - Base64 encoder (UTF-8).
   * @returns {{summary: function(): Object, positions: function(): Array<Object>,
   *   transactionsPage: function(?string): {items: Array<Object>, nextPagePath: ?string},
   *   dividendsPage: function(?string): {items: Array<Object>, nextPagePath: ?string}}} Client.
   */
  function create(http, credentials, base64) {
    const authorization = 'Basic ' + base64(credentials.key + ':' + credentials.secret);

    function get(path) {
      return http.request({ method: 'GET', url: urlFor(path), headers: { Authorization: authorization, Accept: 'application/json' } });
    }

    /**
     * @param {string} firstPath - Endpoint path of the first page.
     * @returns {function(?string): {items: Array<Object>, nextPagePath: ?string}} Reader of one page.
     */
    function pages(firstPath) {
      return function (nextPagePath) {
        const page = get(nextPagePath || firstPath + '?limit=' + T212.PAGE_LIMIT) || {};
        return { items: page.items || [], nextPagePath: page.nextPagePath || null };
      };
    }

    return {
      /** @returns {Object} Account summary: id, currency, cash, investments, totalValue (scope `account`). */
      summary: function () { return get('/equity/account/summary'); },
      /** @returns {Array<Object>} All open positions (scope `portfolio`). */
      positions: function () { return get('/equity/positions') || []; },
      /** Cash movements, newest first (scope `history:transactions`). */
      transactionsPage: pages('/equity/history/transactions'),
      /** Paid dividends, newest first (scope `history:dividends`). */
      dividendsPage: pages('/equity/history/dividends')
    };
  }

  ns.T212Client = { create: create, urlFor: urlFor };
})(globalThis.BankSync);
