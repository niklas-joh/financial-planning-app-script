/**
 * @fileoverview Pluxee connections come from Script Properties, one pair per holder:
 * `PLUXEE_CARD_<HOLDER>` (the card id, e.g. `LUX000...`, as in the portal's transactions URL) and
 * `PLUXEE_REFRESH_TOKEN_<HOLDER>` (copied once from the consumer portal after logging in). The
 * refresh token is replaced here whenever PluxeeConnect hands out a new one, so the owner pastes it
 * again only after the login has lapsed. Token values never leave this module and the client.
 */
(function (ns) {
  const PLUXEE = ns.Config.PLUXEE;
  const KEYS = ns.Config.PROPERTY_KEYS;
  const PREFIXES = [KEYS.PLUXEE_CARD_PREFIX, KEYS.PLUXEE_REFRESH_TOKEN_PREFIX];

  /**
   * Holder suffixes that have a card or a refresh-token property.
   * @param {Array<string>} keys - All Script Property names.
   * @returns {Array<string>} Suffixes, sorted, e.g. ['NIKLAS'].
   */
  function suffixes(keys) {
    const found = {};
    (keys || []).forEach(function (k) {
      PREFIXES.forEach(function (prefix) {
        if (k.indexOf(prefix) === 0 && k.length > prefix.length) found[k.slice(prefix.length)] = true;
      });
    });
    return Object.keys(found).sort();
  }

  /**
   * @param {string} suffix - Holder suffix of the property names.
   * @returns {import('../../core/config.js').ConnectionConfig} Connection keyed 'pluxee-<holder>'.
   */
  function connectionFor(suffix) {
    const holder = ns.Connections.slug(suffix);
    return {
      key: PLUXEE.PROVIDER_ID + '-' + holder,
      holder: holder,
      provider: PLUXEE.PROVIDER_ID,
      propertySuffix: suffix,
      bank: { aspspName: PLUXEE.BANK_NAME, country: 'LU' }
    };
  }

  /**
   * @param {import('../../core/config.js').ConnectionConfig} connection - Pluxee connection.
   * @returns {{card: string, token: string}} Property names of its card id and its refresh token.
   */
  function names(connection) {
    return {
      card: KEYS.PLUXEE_CARD_PREFIX + connection.propertySuffix,
      token: KEYS.PLUXEE_REFRESH_TOKEN_PREFIX + connection.propertySuffix
    };
  }

  /**
   * @param {{getKeys: function(): Array<string>}} props - Script Properties.
   * @returns {Array<import('../../core/config.js').ConnectionConfig>} One connection per holder with either property.
   */
  function connections(props) {
    return suffixes(props.getKeys()).map(connectionFor);
  }

  /**
   * Property names for a holder chosen at connect time, deriving the suffix the same way (upper case) so the
   * Connect Pluxee flow writes exactly the names the reader later scans for.
   * @param {string} holder - Holder, e.g. 'niklas'.
   * @returns {{card: string, token: string}} The card and refresh-token property names.
   */
  function namesForHolder(holder) {
    const suffix = String(holder).toUpperCase();
    return { card: KEYS.PLUXEE_CARD_PREFIX + suffix, token: KEYS.PLUXEE_REFRESH_TOKEN_PREFIX + suffix };
  }

  /**
   * Saves a holder's refresh token (the Connect Pluxee flow). The value is a secret: never log or show it.
   * @param {{setProperty: function(string, string): void}} props - Script Properties.
   * @param {string} holder - Holder, e.g. 'niklas'.
   * @param {string} token - The refresh token from PluxeeConnect.
   */
  function saveToken(props, holder, token) {
    props.setProperty(namesForHolder(holder).token, token);
  }

  /**
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Pluxee connection.
   * @returns {boolean} True when both the card id and the refresh token are set.
   */
  function isComplete(props, connection) {
    const n = names(connection);
    return Boolean(props.getProperty(n.card) && props.getProperty(n.token));
  }

  /**
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Pluxee connection.
   * @returns {string} The card id, trimmed.
   * @throws {Error} ConfigError when missing or not letters and digits only (it goes into a URL path).
   */
  function cardId(props, connection) {
    const name = names(connection).card;
    const value = String(props.getProperty(name) || '').trim();
    if (!value) throw ns.Errors.config('Set ' + name + ' in Script Properties');
    if (!/^[A-Za-z0-9]+$/.test(value)) throw ns.Errors.config(name + ' must be letters and digits only (e.g. LUX000...)');
    return value;
  }

  /**
   * Store of the connection's refresh token, for the client only. Callers must never log or show it.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): void}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Pluxee connection.
   * @returns {{name: string, load: function(): string, save: function(string): void}} Token store; `name` is the
   *   property's name (for messages), never its value.
   */
  function tokenStore(props, connection) {
    const name = names(connection).token;
    return {
      name: name,
      load: function () {
        const value = String(props.getProperty(name) || '').trim();
        if (!value) throw ns.Errors.config('Set ' + name + ' in Script Properties');
        return value;
      },
      save: function (token) { props.setProperty(name, token); }
    };
  }

  /**
   * Status lines for Show status: names and set/MISSING only.
   * @param {{getKeys: function(): Array<string>, getProperty: function(string): ?string}} props - Script Properties.
   * @returns {Array<string>} One line per holder, or one line saying none is set up.
   */
  function statusLines(props) {
    const all = connections(props);
    if (!all.length) return [PLUXEE.BANK_NAME + ': not set up (optional; see the README)'];
    return all.map(function (c) {
      const n = names(c);
      const states = [n.card, n.token].map(function (name) { return name + ': ' + (props.getProperty(name) ? 'set' : 'MISSING'); });
      return ns.Connections.label(c) + ' - ' + states.join('; ');
    });
  }

  ns.PluxeeCredentials = {
    connections: connections, isComplete: isComplete, cardId: cardId, tokenStore: tokenStore, statusLines: statusLines,
    namesForHolder: namesForHolder, saveToken: saveToken
  };
})(globalThis.BankSync);
