/**
 * @fileoverview Composition root: wires Apps Script services into the modules.
 */
(function (ns) {
  /**
   * Builds the shared object graph (Enable Banking client, sessions and connections, Script Properties, clock).
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
   * Apps Script helpers a provider may need (Trading 212: HTTP with its rate-limit hook, Base64, local dates;
   * CSV import: its Drive folder).
   * @param {function(): number} [msLeft] - Time left in the current sync run.
   * @returns {import('../core/provider-registry.js').ProviderEnv} Helpers.
   */
  function providerEnv(msLeft) {
    const formatDate = ns.Gas.dateFormatter(ns.Config.TIMEZONE);
    return {
      http: ns.Gas.http,
      base64: ns.Gas.base64,
      toLocalDate: function (iso) { return formatDate(new Date(iso)); },
      msLeft: msLeft,
      driveFolder: ns.Gas.driveFolder
    };
  }

  /**
   * Builds a provider for one connection.
   * @param {import('../core/config.js').ConnectionConfig} connection - Login or API key to serve.
   * @param {Object} [wired] - Services from `services()`, to avoid wiring twice.
   * @param {import('../core/provider-registry.js').ProviderEnv} [env] - Helpers; built when omitted.
   * @returns {import('../providers/bank-provider.js').BankProvider} Provider.
   */
  function providerFor(connection, wired, env) {
    return ns.ProviderRegistry.create(connection, wired || services(), env || providerEnv());
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
   * Builds the store over the three transaction tabs (`transactions`, `investment_transactions`,
   * `unclassified_transactions`); each row is written to its tab by `TxnRouting.tabFor`.
   * @param {Object} spreadsheet - The Bank Sync spreadsheet.
   * @param {function(Date): string} formatDate - Date to ISO date in the configured time zone.
   * @returns {Object} Routed store.
   */
  function transactionsStore(spreadsheet, formatDate) {
    const stores = {};
    ns.TxnRouting.tabs().forEach(function (tab) { stores[tab] = ns.TransactionsStore.create(spreadsheet, formatDate, tab); });
    return ns.TxnRouting.createStore(stores);
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
    const startedAt = wired.nowMs();
    const env = providerEnv(function () { return startedAt + ns.Config.TRANSACTIONS.TIME_BUDGET_MS - wired.nowMs(); });
    return {
      connections: ns.ProviderRegistry.connections(wired),
      isConnected: function (c) { return ns.ProviderRegistry.isConnected(c, wired); },
      providerFor: function (c) { return providerFor(c, wired, env); },
      store: transactionsStore(spreadsheet, formatDate),
      rawWriter: ns.RawWriter.create(spreadsheet),
      nowMs: wired.nowMs,
      todayIso: todayIso,
      fx: fxService(spreadsheet, formatDate, todayIso)
    };
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
      consents: function () { return ns.ProviderRegistry.consents(wired); },
      send: ns.Gas.mailOwner(),
      props: wired.props,
      nowMs: wired.nowMs,
      nowLabel: function () { return nowLabel(new Date()); }
    };
  }

  ns.Container = {
    services: services, providerFor: providerFor, fxService: fxService,
    syncAllDeps: syncAllDeps, autoSyncDeps: autoSyncDeps
  };
})(globalThis.BankSync);
