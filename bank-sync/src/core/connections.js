/**
 * @fileoverview Connections: one login at one bank (bank plus holder). They are
 * created by the Connect flow and kept in Script Properties, so adding a bank or
 * another person needs no code or config change. Also the pure helpers behind the
 * bank picker.
 */
(function (ns) {
  const MAX_CHOICES = 25;
  const NOT_ENABLED_CODES = ['NO_ACCOUNTS_ADDED', 'ACCESS_DENIED'];

  /**
   * Lowercase, dash-separated form of free text.
   * @param {string} text - Any text.
   * @returns {string} e.g. 'Danske Bank' becomes 'danske-bank'.
   */
  function slug(text) {
    return String(text).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  }

  /**
   * Builds a connection (not stored).
   * @param {{bankName: string, country: string, holder: string}} def - What the user chose.
   * @returns {import('./config.js').ConnectionConfig} Connection keyed by bank plus holder.
   * @throws {Error} InputError if the holder is empty after cleaning.
   */
  function build(def) {
    const holder = slug(def.holder);
    if (!holder) throw ns.Errors.create('InputError', 'Enter a short name for whose login this is');
    return {
      key: slug(def.bankName) + '-' + holder,
      holder: holder,
      bank: { aspspName: def.bankName, country: String(def.country).toUpperCase() }
    };
  }

  /**
   * Human label for menus and alerts.
   * @param {import('./config.js').ConnectionConfig} connection - Connection.
   * @returns {string} e.g. "Revolut (niklas)".
   */
  function label(connection) {
    return connection.bank.aspspName + ' (' + connection.holder + ')';
  }

  /**
   * Finds banks for what the user typed.
   * @param {Array<string>} names - Bank names listed by Enable Banking.
   * @param {string} query - Typed text.
   * @returns {{bank: ?string, matches: Array<string>}} `bank` is set when the answer is unambiguous
   *   (an exact name, or the only match); `matches` lists up to 25 candidates otherwise.
   */
  function matchBanks(names, query) {
    const wanted = String(query).trim().toLowerCase();
    if (!wanted) return { bank: null, matches: [] };
    const exact = names.filter(function (n) { return n.toLowerCase() === wanted; });
    if (exact.length === 1) return { bank: exact[0], matches: exact };
    const partial = names.filter(function (n) { return n.toLowerCase().indexOf(wanted) !== -1; }).sort();
    return { bank: partial.length === 1 ? partial[0] : null, matches: partial.slice(0, MAX_CHOICES) };
  }

  /**
   * Text to show when Enable Banking refuses a bank that has not been enabled, or null if
   * the error is something else.
   * @param {Error} error - Error from starting or completing authorization.
   * @returns {?string} Message telling the user what to do.
   */
  function explainNotEnabled(error) {
    if (!error || NOT_ENABLED_CODES.indexOf(error.apiCode) === -1) return null;
    return 'This bank is not enabled for your Enable Banking application yet. Open the Enable Banking ' +
      'Control Panel, link or whitelist the account for this bank under your application, then run ' +
      '"Connect a bank..." again.';
  }

  /**
   * Creates the connection store.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): *}} props - Script Properties.
   * @returns {{all: function(): Array<Object>, find: function(string): ?Object, save: function(Object): void}} Store.
   */
  function create(props) {
    const KEY = ns.Config.PROPERTY_KEYS.EB_CONNECTIONS;
    const read = function () { const raw = props.getProperty(KEY); return raw ? JSON.parse(raw) : {}; };

    return {
      /** @returns {Array<import('./config.js').ConnectionConfig>} Stored connections, ordered by key. */
      all: function () {
        const map = read();
        return Object.keys(map).sort().map(function (key) { return map[key]; });
      },
      /**
       * @param {string} key - Connection key.
       * @returns {?import('./config.js').ConnectionConfig} The connection, or null.
       */
      find: function (key) { return read()[key] || null; },
      /**
       * Adds or replaces a connection.
       * @param {import('./config.js').ConnectionConfig} connection - Connection to store.
       */
      save: function (connection) {
        const map = read();
        map[connection.key] = connection;
        props.setProperty(KEY, JSON.stringify(map));
      }
    };
  }

  /**
   * Resolves a typed answer ("1", " 2 ") to a connection from a numbered list.
   * @param {Array<import('./config.js').ConnectionConfig>} connections - The numbered list.
   * @param {string} answer - Text the user typed.
   * @returns {import('./config.js').ConnectionConfig} The chosen connection.
   * @throws {Error} InputError if the answer is not a listed number.
   */
  function choose(connections, answer) {
    const text = String(answer).trim();
    const choice = /^\d+$/.test(text) ? connections[Number(text) - 1] : undefined;
    if (!choice) throw ns.Errors.create('InputError', 'Not a listed number');
    return choice;
  }

  ns.Connections = {
    slug: slug, build: build, label: label, matchBanks: matchBanks,
    explainNotEnabled: explainNotEnabled, create: create, choose: choose
  };
})(globalThis.BankSync);
