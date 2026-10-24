/**
 * @fileoverview The only place that touches Apps Script services directly for
 * crypto, HTTP and storage. Everything else receives these as dependencies.
 */
(function (ns) {
  ns.Gas = {
    /** @returns {import('./jwt.js').JwtDeps} Encoder/signer backed by Utilities. */
    jwtDeps: function () {
      return {
        base64UrlEncode: function (input) {
          const encoded = typeof input === 'string'
            ? Utilities.base64EncodeWebSafe(input, Utilities.Charset.UTF_8)
            : Utilities.base64EncodeWebSafe(input);
          return encoded.replace(/=+$/, '');
        },
        signRs256: function (data, privateKey) {
          return Utilities.computeRsaSha256Signature(data, privateKey);
        }
      };
    },

    /** @returns {{request: function(Object): *}} HTTP client backed by UrlFetchApp. */
    http: function () {
      return ns.Http.create({
        fetch: function (url, params) { return UrlFetchApp.fetch(url, params); },
        sleep: function (ms) { Utilities.sleep(ms); }
      });
    },

    /** @returns {Object} Script Properties service. */
    props: function () {
      return PropertiesService.getScriptProperties();
    },

    /** @returns {Object} Script cache service. */
    cache: function () {
      return CacheService.getScriptCache();
    }
  };
})(globalThis.BankSync);
