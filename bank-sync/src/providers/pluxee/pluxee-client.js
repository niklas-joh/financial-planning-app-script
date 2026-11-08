/**
 * @fileoverview Thin client for the Pluxee consumer portal's own API (unofficial; read only).
 * Signing in needs an emailed code and a captcha, so it is done once by the owner in the portal;
 * this client only renews access with the stored refresh token (OpenID Connect `refresh_token`
 * grant against PluxeeConnect, as the portal does) and reads the card's transactions. When the
 * token endpoint returns a new refresh token, it is saved before anything else is called, so a
 * rotated token is never lost. Tokens never leave this module.
 *
 * The API also wants the portal's subscription key (`ocp-apim-subscription-key`), a fixed value
 * built into the portal's script and sent by every visitor's browser. It is read from that script
 * once per connection in each run (two page reads; one card connection is expected), so a key Pluxee
 * rotates is picked up without a code change.
 */
(function (ns) {
  const PLUXEE = ns.Config.PLUXEE;

  /**
   * Whether a token-endpoint failure means the login itself is over (the owner must sign in again). Other
   * refusals (e.g. a 400 because Pluxee changed its client id or parameters) stay HttpErrors with the status,
   * since a new refresh token would not fix them.
   * @param {Error} e - Error from the token request.
   * @returns {boolean} True for OAuth `invalid_grant` and HTTP 401.
   */
  function isLoginExpired(e) {
    return e.apiCode === 'invalid_grant' || e.status === 401;
  }

  /** The key in the portal's script: `backendAPIMKey:` (runtime config) or `VITE_BACKEND_APIM_KEY:` (build env). */
  const KEY_PATTERN = /(?:backendAPIMKey|VITE_BACKEND_APIM_KEY)\s*:\s*[`'"]([0-9a-fA-F]{32})[`'"]/;
  /** The portal's main script in its page, e.g. `/assets/index-DatvBM8E.js`. */
  const SCRIPT_PATTERN = /src="(\/assets\/index-[\w-]+\.js)"/;

  /**
   * Finds the API subscription key in the portal's page and script.
   * @param {string} page - HTML of the portal's start page.
   * @param {function(string): string} getText - Reads a URL as text.
   * @returns {string} The key.
   * @throws {Error} DataError when the page names no script or the script holds no key.
   */
  function subscriptionKey(page, getText) {
    const script = SCRIPT_PATTERN.exec(String(page || ''));
    if (!script) throw ns.Errors.create('DataError', 'Pluxee portal page names no main script (the portal changed)');
    const key = KEY_PATTERN.exec(String(getText(PLUXEE.PORTAL_ORIGIN + script[1]) || ''));
    if (!key) throw ns.Errors.create('DataError', 'Pluxee portal script holds no API subscription key (the portal changed)');
    return key[1];
  }

  /**
   * The OAuth token endpoint from PluxeeConnect's discovery document (shared by the refresh grant here and the
   * sign-in code exchange in `PluxeeConnect`).
   * @param {{request: function(Object): *}} http - JSON HTTP client.
   * @returns {string} The token endpoint.
   * @throws {Error} DataError when the document names none on PluxeeConnect's own host.
   */
  function tokenEndpoint(http) {
    const doc = http.request({ method: 'GET', url: PLUXEE.ISSUER + '/.well-known/openid-configuration' }) || {};
    const url = String(doc.token_endpoint || '');
    if (url.indexOf(PLUXEE.ISSUER + '/') !== 0) throw ns.Errors.create('DataError', 'PluxeeConnect discovery names no usable token endpoint');
    return url;
  }

  /**
   * Creates a client.
   * @param {{request: function(Object): *}} http - JSON HTTP client.
   * @param {{name: string, load: function(): string, save: function(string): void}} tokens - Refresh token store
   *   (`PluxeeCredentials.tokenStore`); `name` is the Script Property named in the re-login message.
   * @param {{request: function(Object): *}} [tokenHttp] - Client for the token request, built without retries:
   *   a retry after a 5xx that came once PluxeeConnect had already rotated the token would send the used one
   *   (and could revoke the login). Defaults to `http` (tests).
   * @returns {{transactions: function(string): Object}} Client.
   */
  function create(http, tokens, tokenHttp) {
    const tokenClient = tokenHttp || http;
    let accessToken = '';
    let apiKey = '';

    /** @returns {string} The portal's API subscription key, read once by this client (one client per connection and run). */
    function portalKey() {
      if (!apiKey) {
        const getText = function (url) { return http.request({ method: 'GET', url: url, text: true }); };
        apiKey = subscriptionKey(getText(PLUXEE.PORTAL_ORIGIN + '/'), getText);
      }
      return apiKey;
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
        out = tokenClient.request({
          method: 'POST',
          url: tokenEndpoint(http),
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
        const key = portalKey(); // before the refresh: a changed portal must not use up a refresh token
        return http.request({
          method: 'GET',
          url: url,
          headers: {
            Authorization: 'Bearer ' + access(),
            'authorization-version': PLUXEE.AUTHORIZATION_VERSION,
            'ocp-apim-subscription-key': key,
            Accept: 'application/json'
          }
        }) || {};
      }
    };
  }

  ns.PluxeeClient = { create: create, subscriptionKey: subscriptionKey, tokenEndpoint: tokenEndpoint };
})(globalThis.BankSync);
