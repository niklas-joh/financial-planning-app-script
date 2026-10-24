/**
 * @fileoverview Supplies a valid Enable Banking JWT, signing a new one only
 * when the cached token has gone.
 */
(function (ns) {
  const CACHE_KEY = 'eb_jwt';
  const CACHE_MARGIN_SECONDS = 300;

  /**
   * @typedef {Object} TokenDeps
   * @property {{getProperty: function(string): ?string}} props - Script Properties holding credentials.
   * @property {{get: function(string): ?string, put: function(string, string, number): *}} cache - Script cache.
   * @property {function(): number} nowSeconds - Clock in Unix seconds.
   * @property {import('../../lib/jwt.js').JwtDeps} jwtDeps - Encoder and signer.
   */

  /**
   * Creates a token provider.
   * @param {TokenDeps} deps - Injected services.
   * @returns {{getToken: function(): string}} Provider.
   * @throws {Error} ConfigError from `getToken` if credentials are not set.
   */
  function create(deps) {
    const cfg = ns.Config.ENABLE_BANKING;
    const keys = ns.Config.PROPERTY_KEYS;

    function readCredential(key) {
      const value = deps.props.getProperty(key);
      if (!value) throw ns.Errors.config('Script Property ' + key + ' is not set');
      return value;
    }

    return {
      getToken: function () {
        const cached = deps.cache.get(CACHE_KEY);
        if (cached) return cached;
        const token = ns.Jwt.create({
          appId: readCredential(keys.EB_APP_ID),
          privateKey: readCredential(keys.EB_PRIVATE_KEY),
          nowSeconds: deps.nowSeconds(),
          ttlSeconds: cfg.JWT_TTL_SECONDS,
          issuer: cfg.JWT_ISSUER,
          audience: cfg.JWT_AUDIENCE
        }, deps.jwtDeps);
        deps.cache.put(CACHE_KEY, token, cfg.JWT_TTL_SECONDS - CACHE_MARGIN_SECONDS);
        return token;
      }
    };
  }

  ns.EbToken = { create: create };
})(globalThis.BankSync);
