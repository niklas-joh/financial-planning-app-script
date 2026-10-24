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

  ns.Connections = { all: all, label: label };
})(globalThis.BankSync);
