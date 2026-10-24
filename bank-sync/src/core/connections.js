/**
 * @fileoverview Read access to the configured connections (one login at one bank).
 */
(function (ns) {
  /**
   * All configured connections, in config order.
   * @returns {Array<import('./config.js').ConnectionConfig>} Connections.
   */
  function all() {
    return Object.keys(ns.Config.CONNECTIONS).map(function (name) { return ns.Config.CONNECTIONS[name]; });
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

  ns.Connections = { all: all, label: label, choose: choose };
})(globalThis.BankSync);
