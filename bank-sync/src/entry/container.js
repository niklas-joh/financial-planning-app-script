/**
 * @fileoverview Composition root: wires Apps Script services into the modules.
 */
(function (ns) {
  /**
   * Builds the object graph for Enable Banking. Sessions stored under an older
   * key are moved to their connection key (idempotent, cheap).
   * @returns {{store: Object, auth: Object, client: Object, props: Object, nowMs: function(): number}} Wired services.
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
    ns.Connections.all().forEach(function (connection) {
      if (connection.legacySessionKey) store.migrateKey(connection.legacySessionKey, connection.key);
    });
    const auth = ns.EbAuth.create({
      client: client,
      store: store,
      getRedirectUrl: function () { return props.getProperty(ns.Config.PROPERTY_KEYS.EB_REDIRECT_URL); },
      nowMs: nowMs,
      uuid: function () { return Utilities.getUuid(); }
    });
    return { store: store, auth: auth, client: client, props: props, nowMs: nowMs };
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
