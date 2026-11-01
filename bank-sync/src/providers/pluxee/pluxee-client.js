/**
 * @fileoverview Thin client for the Pluxee consumer portal's own API (unofficial; read only).
 * Signing in needs an emailed code and a captcha, so it is done once by the owner in the portal;
 * this client only renews access with the stored refresh token (OpenID Connect `refresh_token`
 * grant against PluxeeConnect, as the portal does) and reads the card's transactions. When the
 * token endpoint returns a new refresh token, it is saved before anything else is called, so a
 * rotated token is never lost. Tokens never leave this module.
 */
(function (ns) {
  const PLUXEE = ns.Config.PLUXEE;

  /**
   * Whether a token-endpoint failure means the login itself is over (the owner must sign in again).
   * @param {Error} e - Error from the token request.
   * @returns {boolean} True for OAuth `invalid_grant` and HTTP 400/401.
   */
  function isLoginExpired(e) {
    return e.apiCode === 'invalid_grant' || e.status === 400 || e.status === 401;
  }

  /**
   * Creates a client.
   * @param {{request: function(Object): *}} http - JSON HTTP client.
   * @param {{name: string, load: function(): string, save: function(string): void}} tokens - Refresh token store
   *   (`PluxeeCredentials.tokenStore`); `name` is the Script Property named in the re-login message.
   * @returns {{transactions: function(string): Object}} Client.
   */
  function create(http, tokens) {
    let accessToken = '';

    /**
     * @returns {string} The token endpoint from PluxeeConnect's discovery document.
     * @throws {Error} DataError when the document names none on PluxeeConnect's own host.
     */
    function tokenEndpoint() {
      const doc = http.request({ method: 'GET', url: PLUXEE.ISSUER + '/.well-known/openid-configuration' }) || {};
      const url = String(doc.token_endpoint || '');
      if (url.indexOf(PLUXEE.ISSUER + '/') !== 0) throw ns.Errors.create('DataError', 'PluxeeConnect discovery names no usable token endpoint');
      return url;
    }

    /**
     * A fresh access token for this run (renewed once, then reused).
     * @returns {string} Access token.
     * @throws {Error} AuthError when the login has lapsed; HttpError for other failures.
     */
    function access() {
      if (accessToken) return accessToken;
      let out;
      try {
        out = http.request({
          method: 'POST',
          url: tokenEndpoint(),
          headers: { Accept: 'application/json' },
          form: { grant_type: 'refresh_token', refresh_token: tokens.load(), client_id: PLUXEE.CLIENT_ID }
        }) || {};
      } catch (e) {
        if (e.name === 'HttpError' && isLoginExpired(e)) {
          throw ns.Errors.auth('Pluxee login has expired: sign in at the consumer portal again and paste the new refresh token into ' + tokens.name);
        }
        throw e;
      }
      if (out.refresh_token) tokens.save(String(out.refresh_token));
      if (!out.access_token) throw ns.Errors.create('DataError', 'PluxeeConnect returned no access token');
      accessToken = String(out.access_token);
      return accessToken;
    }

    return {
      /**
       * The card and its latest transactions, newest first.
       * @param {string} cardId - Card id (letters and digits; checked by `PluxeeCredentials.cardId`).
       * @returns {{card: Object, transactions: Array<Object>}} The portal's response.
       */
      transactions: function (cardId) {
        const url = PLUXEE.API_ORIGIN + PLUXEE.TRANSACTIONS_PATH + encodeURIComponent(cardId) + '/transactions?limit=' + PLUXEE.PAGE_LIMIT;
        return http.request({ method: 'GET', url: url, headers: { Authorization: 'Bearer ' + access(), Accept: 'application/json' } }) || {};
      }
    };
  }

  ns.PluxeeClient = { create: create };
})(globalThis.BankSync);
