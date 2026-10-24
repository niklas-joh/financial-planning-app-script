/**
 * @fileoverview Thin Enable Banking REST client: one method per endpoint,
 * no business logic.
 */
(function (ns) {
  /**
   * Creates a client.
   * @param {{request: function(Object): *}} http - JSON HTTP client.
   * @param {function(): string} getToken - Returns a valid JWT.
   * @returns {Object} Client with listAspsps, startAuth, createSession, getSession, getBalances.
   */
  function create(http, getToken) {
    const base = ns.Config.ENABLE_BANKING.BASE_URL;

    function call(method, path, body) {
      return http.request({
        method: method,
        url: base + path,
        headers: { Authorization: 'Bearer ' + getToken(), Accept: 'application/json' },
        body: body
      });
    }

    return {
      /** @param {string} country - ISO 3166 code. @returns {{aspsps: Array<Object>}} Banks. */
      listAspsps: function (country) {
        return call('GET', '/aspsps?country=' + encodeURIComponent(country));
      },
      /** @param {Object} body - POST /auth body. @returns {{url: string}} Login URL. */
      startAuth: function (body) {
        return call('POST', '/auth', body);
      },
      /** @param {string} code - Authorization code from the redirect. @returns {Object} Session. */
      createSession: function (code) {
        return call('POST', '/sessions', { code: code });
      },
      /** @param {string} sessionId - Session id. @returns {Object} Session state. */
      getSession: function (sessionId) {
        return call('GET', '/sessions/' + encodeURIComponent(sessionId));
      },
      /** @param {string} accountUid - Account uid. @returns {{balances: Array<Object>}} Balances. */
      getBalances: function (accountUid) {
        return call('GET', '/accounts/' + encodeURIComponent(accountUid) + '/balances');
      }
    };
  }

  ns.EbClient = { create: create };
})(globalThis.BankSync);
