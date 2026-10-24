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

  var PEM_PATTERN = /-----BEGIN ([A-Z ]+)-----([\s\S]*?)-----END \1-----/;

  /**
   * Rebuilds a PKCS#8 PEM pasted into a Script Property. The property editor is
   * a single-line field, so pasted line breaks arrive as spaces (or as literal
   * "\n"); Utilities.computeRsaSha256Signature then rejects the key. The base64
   * body is stripped of all whitespace and re-wrapped at 64 characters.
   * Error messages never include key material.
   * @param {string} pem - Raw PEM text.
   * @returns {string} Well-formed PKCS#8 PEM.
   * @throws {Error} If the text has no BEGIN/END block, or the key is PKCS#1.
   */
  function normalizePem(pem) {
    var text = String(pem).replace(/\\n/g, '\n');
    var match = PEM_PATTERN.exec(text);
    if (!match) {
      throw new Error('EB_PRIVATE_KEY is not a PEM key: paste the whole .pem, including the BEGIN and END lines.');
    }
    if (match[1] !== 'PRIVATE KEY') {
      throw new Error('EB_PRIVATE_KEY must be a PKCS#8 key ("BEGIN PRIVATE KEY"), got "BEGIN ' + match[1] +
        '". Convert it with: openssl pkcs8 -topk8 -nocrypt -in key.pem -out key-pkcs8.pem');
    }
    var body = match[2].replace(/\s+/g, '');
    return '-----BEGIN PRIVATE KEY-----\n' + body.match(/.{1,64}/g).join('\n') + '\n-----END PRIVATE KEY-----\n';
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
