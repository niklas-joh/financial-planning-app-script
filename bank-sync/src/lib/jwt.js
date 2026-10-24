/**
 * @fileoverview RS256 JWT creation for Enable Banking. Pure: crypto and
 * encoding are injected, so it runs under Node as well as Apps Script.
 */
(function (ns) {
  /**
   * @typedef {Object} JwtDeps
   * @property {function((string|Array<number>)): string} base64UrlEncode - Unpadded base64url.
   * @property {function(string, string): Array<number>} signRs256 - Signs data with a PEM key.
   */

  /**
   * @typedef {Object} JwtParams
   * @property {string} appId - Application ID, used as the `kid` header.
   * @property {string} privateKey - PKCS#8 PEM private key.
   * @property {number} nowSeconds - Current Unix time in seconds.
   * @property {number} ttlSeconds - Token lifetime in seconds.
   * @property {string} issuer - `iss` claim.
   * @property {string} audience - `aud` claim.
   */

  /**
   * Normalizes a PEM pasted into a property (literal "\n" or stray whitespace).
   * @param {string} pem - Raw PEM text.
   * @returns {string} PEM with real newlines and no surrounding whitespace.
   */
  function normalizePem(pem) {
    return String(pem).replace(/\\n/g, '\n').trim();
  }

  /**
   * Builds and signs a JWT.
   * @param {JwtParams} params - Header/claim inputs and key.
   * @param {JwtDeps} deps - Injected encoder and signer.
   * @returns {string} Compact serialized JWT.
   */
  function create(params, deps) {
    const header = { typ: 'JWT', alg: 'RS256', kid: params.appId };
    const claims = {
      iss: params.issuer,
      aud: params.audience,
      iat: params.nowSeconds,
      exp: params.nowSeconds + params.ttlSeconds
    };
    const signingInput =
      deps.base64UrlEncode(JSON.stringify(header)) + '.' +
      deps.base64UrlEncode(JSON.stringify(claims));
    const signature = deps.signRs256(signingInput, normalizePem(params.privateKey));
    return signingInput + '.' + deps.base64UrlEncode(signature);
  }

  ns.Jwt = { create: create, normalizePem: normalizePem };
})(globalThis.BankSync);
