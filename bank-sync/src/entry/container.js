/**
 * @fileoverview Composition root: wires Apps Script services into the modules.
 */
(function (ns) {
  /**
   * One-time move of the pre-connections session (stored under the bank key) to its
   * connection key, creating the connection record. Guarded by a property flag so it
   * does not touch Script Properties again.
   * @param {Object} props - Script Properties.
   * @param {Object} store - EbSessionStore.
   * @param {{save: function(Object): void}} connections - Connection store.
   */
  function migrateLegacy(props, store, connections) {
    const flag = ns.Config.PROPERTY_KEYS.EB_MIGRATED;
    if (props.getProperty(flag)) return;
    const legacy = ns.Config.LEGACY_CONNECTION;
    const connection = ns.Connections.build(legacy);
    if (store.migrateKey(legacy.sessionKey, connection.key)) connections.save(connection);
    props.setProperty(flag, '1');
  }

  /**
   * Builds the object graph for Enable Banking.
   * @returns {{store: Object, auth: Object, client: Object, props: Object, connections: Object,
   *   nowMs: function(): number}} Wired services.
   */
  function services() {
    const props = ns.Gas.props();
    const nowMs = function () { return Date.now(); };
    const token = ns.EbToken.create({
      props: props,
      cache: ns.Gas.cache(),
      nowSeconds: function () { return Math.floor(nowMs() / 1000); },
      jwtDeps: ns.Gas.jwtDeps()
    });
    const client = ns.EbClient.create(ns.Gas.http(), token.getToken);
    const store = ns.EbSessionStore.create(props, nowMs);
    const connections = ns.Connections.create(props);
    migrateLegacy(props, store, connections);
    const auth = ns.EbAuth.create({
      client: client,
      store: store,
      connections: connections,
      getRedirectUrl: function () { return props.getProperty(ns.Config.PROPERTY_KEYS.EB_REDIRECT_URL); },
      nowMs: nowMs,
      uuid: function () { return Utilities.getUuid(); }
    });
    return { store: store, auth: auth, client: client, props: props, connections: connections, nowMs: nowMs };
  }

  /**
   * Builds a provider for one connection.
   * @param {import('../core/config.js').ConnectionConfig} connection - Login to serve.
   * @param {Object} [wired] - Services from `services()`, to avoid wiring twice.
   * @returns {import('../providers/bank-provider.js').BankProvider} Provider.
   */
  function providerFor(connection, wired) {
    const s = wired || services();
    return ns.EnableBankingProvider.create({
      connection: connection, client: s.client, store: s.store, auth: s.auth, nowMs: s.nowMs
    });
  }

  ns.Container = { services: services, providerFor: providerFor };
})(globalThis.BankSync);
