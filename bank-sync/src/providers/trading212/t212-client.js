/**
 * @fileoverview Thin Trading 212 public API client: one method per endpoint, no
 * business logic. Read-only endpoints, plus the request for a CSV report (it creates a report
 * of the account's history, nothing else). Auth is HTTP Basic with the API key and secret; the
 * header is built here and never leaves this module.
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
   * Whether a report download link is on Trading 212's own servers, the only place the key may be sent.
   * Only `https://<host>/...` links are accepted; a host followed by anything but a port or the path
   * (e.g. `user@host`) is refused.
   * @param {string} link - Download link from the reports list.
   * @returns {boolean} True for trading212.com and its subdomains, false for any other host.
   * @throws {Error} DataError if the link is not an https URL.
   */
  function isOwnHost(link) {
    const match = /^https:\/\/([a-z0-9.-]+)(?::\d+)?(?:[/?#]|$)/i.exec(String(link || ''));
    if (!match) throw ns.Errors.create('DataError', 'Unexpected Trading 212 report link');
    const host = match[1].toLowerCase();
    return host === 'trading212.com' || /\.trading212\.com$/.test(host);
  }

  /**
   * The URL a redirect points to: an absolute link as is, a path on the same origin, or '' for anything else.
   * @param {string} from - URL that answered with the redirect (https, checked by `isOwnHost`).
   * @param {string} location - Its `Location` header.
   * @returns {string} Next URL, or ''.
   */
  function redirectTarget(from, location) {
    const text = String(location || '');
    if (/^https?:\/\//i.test(text)) return text;
    if (text.charAt(0) === '/' && text.charAt(1) !== '/') return /^https:\/\/[^/?#]+/i.exec(from)[0] + text;
    return '';
  }

  /**
   * Creates a client.
   * @param {{request: function(Object): *}} http - JSON HTTP client (with the rate-limit retry hook).
   * @param {{key: string, secret: string}} credentials - API key and secret.
   * @param {function(string): string} base64 - Base64 encoder (UTF-8).
   * @returns {{summary: function(): Object, positions: function(): Array<Object>,
   *   transactionsPage: function(?string): {items: Array<Object>, nextPagePath: ?string},
   *   dividendsPage: function(?string): {items: Array<Object>, nextPagePath: ?string},
   *   requestReport: function(Object): {reportId: number}, reports: function(): Array<Object>,
   *   download: function(string): string}} Client.
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
      dividendsPage: pages('/equity/history/dividends'),
      /**
       * Asks Trading 212 to build a CSV report (1 per 30 s); it is built in the background.
       * @param {{dataIncluded: Object<string, boolean>, timeFrom: string, timeTo: string}} body - What and which period.
       * @returns {{reportId: number}} The new report's id.
       */
      requestReport: function (body) {
        return http.request({ method: 'POST', url: urlFor('/equity/history/exports'), body: body,
          headers: { Authorization: authorization, Accept: 'application/json' } }) || {};
      },
      /** @returns {Array<Object>} Every requested report with its status and, once finished, `downloadLink` (1 per minute). */
      reports: function () { return get('/equity/history/exports') || []; },
      /**
       * Downloads a finished report. Redirects are followed here, one hop at a time (at most
       * `TRADING212.MAX_REDIRECTS`), and the key is sent only to a hop on Trading 212's own hosts (`isOwnHost`),
       * never to the storage service a link or a redirect may point to.
       * @param {string} link - `downloadLink` from `reports`.
       * @returns {string} The CSV text.
       * @throws {Error} DataError for a link or redirect that is not https, has no target, or redirects too often.
       */
      download: function (link) {
        let url = link;
        for (let hop = 0; hop <= T212.MAX_REDIRECTS; hop++) {
          const out = http.request({
            method: 'GET', url: url, text: true, followRedirects: false,
            headers: isOwnHost(url) ? { Authorization: authorization } : {}
          });
          if (typeof out === 'string') return out;
          url = redirectTarget(url, out && out.redirect);
          if (!url) throw ns.Errors.create('DataError', 'A Trading 212 report redirect has no usable target');
        }
        throw ns.Errors.create('DataError', 'The Trading 212 report link redirects too often');
      }
    };
  }

  ns.T212Client = { create: create, urlFor: urlFor, isOwnHost: isOwnHost };
})(globalThis.BankSync);
