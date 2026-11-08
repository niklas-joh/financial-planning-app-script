/**
 * @fileoverview One-time sign-in for a Pluxee card, through the Pluxee iOS app's OAuth client rather than
 * the consumer portal's. The app client (OpenID Connect, PKCE, a public native client) is granted
 * long-lived refresh tokens, where the portal's web client is refused them and its token lapses in minutes.
 * So this replaces copying the portal's refresh token from developer tools.
 *
 * Signing in still needs a captcha and an emailed code, which a script cannot do, so the owner signs in once
 * in a browser. The flow is two menu steps:
 *   - `start(holder)` builds the authorization URL (with a fresh PKCE verifier kept in the script cache) and
 *     returns it. The owner opens it, signs in, and is redirected to the app's own address
 *     (`app.pluxee.consumers.eva://home/?code=...`), which a browser cannot open — the address, carrying the
 *     one-time `code`, stays in the address bar.
 *   - `finish(pasted)` takes that address, exchanges the `code` (with the kept verifier) for tokens at
 *     PluxeeConnect, and saves the refresh token into `PLUXEE_REFRESH_TOKEN_<HOLDER>`. Each later sync renews
 *     it the same way the card provider already does.
 *
 * The verifier and the `code` are single-use secrets: they are held only in the script cache for the minutes
 * between the two steps, never logged, and cleared once used. Tokens never leave this module and the client.
 */
(function (ns) {
  const PLUXEE = ns.Config.PLUXEE;
  /** Script-cache key holding the in-progress sign-in (holder, PKCE verifier, state) between the two steps. */
  const PENDING_KEY = 'pluxee_connect_pending';
  /** How long a started sign-in waits for the pasted code before the owner must start again (seconds). */
  const PENDING_TTL_S = 1800;

  /**
   * A PKCE code verifier: 64 hex characters (two UUIDs, dashes removed), which satisfies RFC 7636's
   * 43-128 unreserved characters and carries about 240 bits of randomness (two v4 UUIDs, less their fixed
   * version and variant bits).
   * @param {function(): string} uuid - Random UUID source.
   * @returns {string} The verifier.
   */
  function verifierFrom(uuid) {
    return (uuid() + uuid()).replace(/-/g, '');
  }

  /**
   * The PluxeeConnect authorization URL for the app client.
   * @param {string} codeChallenge - base64url(SHA-256(verifier)).
   * @param {string} state - Opaque value echoed back, checked on return.
   * @param {string} nonce - OpenID nonce.
   * @returns {string} The URL to open in a browser.
   */
  function authUrl(codeChallenge, state, nonce) {
    const params = {
      client_id: PLUXEE.CLIENT_ID,
      redirect_uri: PLUXEE.REDIRECT_URI,
      response_type: 'code',
      scope: PLUXEE.SCOPE,
      resource: PLUXEE.RESOURCE,
      code_challenge: codeChallenge,
      code_challenge_method: 'S256',
      nonce: nonce,
      state: state,
      ui_locales: 'en-LU'
    };
    const query = Object.keys(params).map(function (k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    }).join('&');
    return PLUXEE.ISSUER + '/oidc/auth?' + query;
  }

  /**
   * Reads one query parameter from a pasted redirect address (a custom-scheme URL or just its query string).
   * @param {string} text - What the owner pasted.
   * @param {string} name - Parameter name.
   * @returns {string} Its decoded value, or '' when absent.
   */
  function queryParam(text, name) {
    const match = new RegExp('[?&]' + name + '=([^&#\\s]*)').exec(String(text || ''));
    if (!match) return '';
    try {
      return decodeURIComponent(match[1].replace(/\+/g, ' '));
    } catch (e) {
      return match[1];
    }
  }

  /**
   * Creates the connect service.
   * @param {Object} deps - Injected services.
   * @param {{request: function(Object): *}} deps.http - JSON HTTP client (discovery).
   * @param {{request: function(Object): *}} deps.tokenHttp - HTTP client without retries for the code exchange:
   *   a retry after a 5xx that already spent the code would fail, and a resend could trip reuse detection.
   * @param {{put: function(string, string, number): void, get: function(string): ?string,
   *   remove: function(string): void}} deps.cache - Script cache, for the between-steps verifier.
   * @param {function(): string} deps.uuid - Random UUID source.
   * @param {function(string): string} deps.sha256 - base64url(SHA-256(text)) for the PKCE challenge.
   * @param {function(string, string): void} deps.saveToken - Saves a holder's refresh token to Script Properties.
   * @returns {{start: function(string): {holder: string, url: string}, finish: function(string): {holder: string}}}
   */
  function create(deps) {
    return {
      /**
       * Begins a sign-in: stores a fresh PKCE verifier and returns the URL to open.
       * @param {string} holderInput - Whose card this is, e.g. 'niklas'.
       * @returns {{holder: string, url: string}} The normalized holder and the authorization URL.
       * @throws {Error} InputError when the holder is empty after cleaning.
       */
      start: function (holderInput) {
        const holder = ns.Connections.slug(holderInput);
        if (!holder) throw ns.Errors.create('InputError', 'Type a short name for whose card this is, e.g. niklas');
        const verifier = verifierFrom(deps.uuid);
        const state = deps.uuid();
        const nonce = deps.uuid();
        deps.cache.put(PENDING_KEY, JSON.stringify({ holder: holder, verifier: verifier, state: state }), PENDING_TTL_S);
        return { holder: holder, url: authUrl(deps.sha256(verifier), state, nonce) };
      },

      /**
       * Completes a sign-in: exchanges the pasted code for tokens and saves the refresh token.
       * @param {string} pasted - The address the browser landed on (holds `code` and `state`).
       * @returns {{holder: string}} Whose refresh token was saved.
       * @throws {Error} InputError/AuthError when nothing is in progress, the paste has no code, the state
       *   does not match, or Pluxee refused the sign-in; DataError when no refresh token comes back.
       */
      finish: function (pasted) {
        const raw = deps.cache.get(PENDING_KEY);
        if (!raw) {
          throw ns.Errors.create('InputError', 'No Pluxee sign-in is in progress, or it has expired. Run Connect Pluxee first.');
        }
        const pending = JSON.parse(raw);
        const refused = queryParam(pasted, 'error');
        if (refused) {
          throw ns.Errors.auth('Pluxee refused the sign-in (' + refused + '). Run Connect Pluxee and try again.');
        }
        const code = queryParam(pasted, 'code');
        if (!code) {
          throw ns.Errors.create('InputError', 'That has no sign-in code in it. Paste the whole address the browser landed on.');
        }
        const state = queryParam(pasted, 'state');
        if (state !== pending.state) { // a missing state fails too: it is the only guard before the code is spent
          throw ns.Errors.create('InputError',
            'This code is from a different sign-in, or the pasted address is incomplete. Run Connect Pluxee again and paste the whole address.');
        }
        const endpoint = ns.PluxeeClient.tokenEndpoint(deps.http);
        let out;
        try {
          out = deps.tokenHttp.request({
            method: 'POST',
            url: endpoint,
            headers: { Accept: 'application/json' },
            form: {
              grant_type: 'authorization_code',
              code: code,
              redirect_uri: PLUXEE.REDIRECT_URI,
              client_id: PLUXEE.CLIENT_ID,
              code_verifier: pending.verifier
            }
          }) || {};
        } finally {
          deps.cache.remove(PENDING_KEY); // the code is single-use whatever the outcome
        }
        if (!out.refresh_token) {
          throw ns.Errors.create('DataError', 'Pluxee returned no refresh token. Run Connect Pluxee and try again.');
        }
        deps.saveToken(pending.holder, String(out.refresh_token));
        return { holder: pending.holder };
      }
    };
  }

  ns.PluxeeConnect = { create: create, authUrl: authUrl, queryParam: queryParam, verifierFrom: verifierFrom };
})(globalThis.BankSync);
