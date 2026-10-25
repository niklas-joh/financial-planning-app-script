/**
 * @fileoverview One-time move of the session that existed before connections (stored under
 * the bank key) to its connection key. Safe to repeat and safe if two callers race: every step
 * writes the same data, so concurrent first calls converge on the same state. Remove this
 * file, `Config.LEGACY_CONNECTION` and the single-map fallback in `EbSessionStore` once the
 * live session has been migrated.
 */
(function (ns) {
  /**
   * Migrates the legacy session and records its connection.
   * @param {Object} props - Script Properties.
   * @param {Object} store - EbSessionStore.
   * @param {{save: function(Object): void}} connections - Connection store.
   * @returns {void}
   */
  function run(props, store, connections) {
    const flag = ns.Config.PROPERTY_KEYS.EB_MIGRATED;
    if (props.getProperty(flag)) return;
    const legacy = ns.Config.LEGACY_CONNECTION;
    const connection = ns.Connections.build(legacy);
    store.migrateKey(legacy.sessionKey, connection.key);
    // Saved whenever the session exists under its new key, not only when this call moved it,
    // so a run that stopped between the two steps is repaired by the next one.
    if (store.loadSession(connection.key)) connections.save(connection);
    props.setProperty(flag, '1');
  }

  ns.LegacyMigration = { run: run };
})(globalThis.BankSync);
