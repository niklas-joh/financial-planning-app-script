/**
 * @fileoverview Composition root: wires Apps Script services into the modules.
 */
(function (ns) {
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
    ns.LegacyMigration.run(props, store, connections);
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

  /**
   * Builds the FX service for one run: ECB rates from Frankfurter, kept on the Rates tab
   * and in the script cache.
   * @param {Object} spreadsheet - The Bank Sync spreadsheet.
   * @param {function(Date): string} formatDate - Date to ISO date in the configured time zone.
   * @param {function(): string} todayIso - Today's date.
   * @returns {Object} FxService.
   */
  function fxService(spreadsheet, formatDate, todayIso) {
    return ns.FxService.create({
      fetchRates: ns.Frankfurter.create(ns.Gas.http()).fetchRates,
      rates: ns.RatesStore.create(spreadsheet, formatDate),
      cache: ns.Gas.cache(),
      todayIso: todayIso
    });
  }

  ns.Container = { services: services, providerFor: providerFor, fxService: fxService };
})(globalThis.BankSync);
