/**
 * @fileoverview Trading 212 connections come from Script Properties: one API key and secret
 * pair per holder, `T212_API_KEY_<HOLDER>` and `T212_API_SECRET_<HOLDER>` (e.g. `_NIKLAS`,
 * `_NIKLAS_ISA`), pasted by the owner. Nothing is stored besides the pair, and values never
 * leave `credentials` (only the client uses them).
 */
(function (ns) {
  const T212 = ns.Config.TRADING212;
  const KEYS = ns.Config.PROPERTY_KEYS;

  /**
   * Holder suffixes that have a key or a secret property.
   * @param {Array<string>} keys - All Script Property names.
   * @returns {Array<string>} Suffixes, sorted, e.g. ['NIKLAS'].
   */
  function suffixes(keys) {
    const found = {};
    (keys || []).forEach(function (k) {
      [KEYS.T212_API_KEY_PREFIX, KEYS.T212_API_SECRET_PREFIX].forEach(function (prefix) {
        if (k.indexOf(prefix) === 0 && k.length > prefix.length) found[k.slice(prefix.length)] = true;
      });
    });
    return Object.keys(found).sort();
  }

  /**
   * @param {string} suffix - Holder suffix of the property names.
   * @returns {import('../../core/config.js').ConnectionConfig} Connection keyed 'trading212-<holder>'.
   */
  function connectionFor(suffix) {
    const holder = ns.Connections.slug(suffix);
    return {
      key: T212.PROVIDER_ID + '-' + holder,
      holder: holder,
      provider: T212.PROVIDER_ID,
      propertySuffix: suffix,
      bank: { aspspName: T212.BANK_NAME, country: '' }
    };
  }

  /**
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {Array<string>} Property names of its key and its secret, in that order.
   */
  function propertyNames(connection) {
    return [KEYS.T212_API_KEY_PREFIX, KEYS.T212_API_SECRET_PREFIX].map(function (prefix) { return prefix + connection.propertySuffix; });
  }

  /**
   * @param {{getKeys: function(): Array<string>}} props - Script Properties.
   * @returns {Array<import('../../core/config.js').ConnectionConfig>} One connection per holder with a key or secret set.
   */
  function connections(props) {
    return suffixes(props.getKeys()).map(connectionFor);
  }

  /**
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {boolean} True when both the key and the secret are set.
   */
  function isComplete(props, connection) {
    return propertyNames(connection).every(function (name) { return Boolean(props.getProperty(name)); });
  }

  /**
   * Reads the pair for the client. Callers must never log or show the result.
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {{key: string, secret: string}} Key and secret.
   * @throws {Error} ConfigError naming the missing property (never its value).
   */
  function credentials(props, connection) {
    const values = propertyNames(connection).map(function (name) {
      const value = props.getProperty(name);
      if (!value) throw ns.Errors.config('Set ' + name + ' in Script Properties');
      return value;
    });
    return { key: values[0], secret: values[1] };
  }

  /**
   * Status lines for Show status: names and set/MISSING only.
   * @param {{getKeys: function(): Array<string>, getProperty: function(string): ?string}} props - Script Properties.
   * @returns {Array<string>} One line per holder, or one line saying no key is set.
   */
  function statusLines(props) {
    const all = connections(props);
    if (!all.length) return [T212.BANK_NAME + ': no API key set (optional; see the README)'];
    return all.map(function (c) {
      const states = propertyNames(c).map(function (name) { return name + ': ' + (props.getProperty(name) ? 'set' : 'MISSING'); });
      return ns.Connections.label(c) + ' - ' + states.join('; ');
    });
  }

  ns.T212Credentials = {
    connections: connections, isComplete: isComplete, credentials: credentials, statusLines: statusLines
  };
})(globalThis.BankSync);
