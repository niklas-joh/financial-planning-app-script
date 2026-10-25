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

  /**
   * Builds everything one Sync all run needs. Used by the menu and by the daily trigger, so
   * both run exactly the same sync.
   * @param {Object} wired - Services from `services()`.
   * @returns {import('../core/sync-all.js').SyncAllDeps} Dependencies for `SyncAll.run`.
   */
  function syncAllDeps(wired) {
    const formatDate = ns.Gas.dateFormatter(ns.Config.TIMEZONE);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const todayIso = function () { return formatDate(new Date()); };
    return {
      connections: wired.connections.all(),
      isConnected: function (c) { return Boolean(wired.store.loadSession(c.key)); },
      providerFor: function (c) { return providerFor(c, wired); },
      store: ns.TransactionsStore.create(spreadsheet, formatDate),
      rawWriter: ns.RawWriter.create(spreadsheet),
      nowMs: wired.nowMs,
      todayIso: todayIso,
      fx: fxService(spreadsheet, formatDate, todayIso)
    };
  }

  /**
   * @param {Object} wired - Services from `services()`.
   * @returns {Array<{connection: Object, validUntil: string}>} Every connection with a session and its consent expiry.
   */
  function consents(wired) {
    return wired.connections.all().map(function (c) {
      const session = wired.store.loadSession(c.key);
      return session ? { connection: c, validUntil: session.accessValidUntil } : null;
    }).filter(Boolean);
  }

  /**
   * Builds the daily-sync services. Sync all runs under the script lock.
   * @returns {import('../core/auto-sync.js').AutoSyncDeps} Dependencies for `AutoSync.runTriggered`.
   */
  function autoSyncDeps() {
    const wired = services();
    const nowLabel = ns.Gas.dateTimeFormatter(ns.Config.TIMEZONE);
    return {
      triggers: ns.Gas.triggers(),
      lastRun: ns.LastRun.create(wired.props),
      sync: function () {
        return ns.Lock.run(LockService, ns.Config.TRANSACTIONS.LOCK_WAIT_MS, function () {
          return ns.SyncAll.run(syncAllDeps(wired));
        });
      },
      consents: function () { return consents(wired); },
      send: ns.Gas.mailOwner(),
      props: wired.props,
      nowMs: wired.nowMs,
      nowLabel: function () { return nowLabel(new Date()); }
    };
  }

  ns.Container = {
    services: services, providerFor: providerFor, fxService: fxService,
    syncAllDeps: syncAllDeps, consents: consents, autoSyncDeps: autoSyncDeps
  };
})(globalThis.BankSync);
