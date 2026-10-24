/**
 * @fileoverview Composition root: wires Apps Script services into the modules.
 */
(function (ns) {
  /**
   * Builds the object graph for Enable Banking.
   * @returns {{store: Object, auth: Object, client: Object, props: Object}} Wired services.
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
   * Builds a provider for one configured bank.
   * @param {import('../core/config.js').BankConfig} bank - Bank to serve.
   * @returns {import('../providers/bank-provider.js').BankProvider} Provider.
   */
  function providerFor(bank) {
    const s = services();
    return ns.EnableBankingProvider.create({
      bank: bank, client: s.client, store: s.store, auth: s.auth, nowMs: s.nowMs
    });
  }

  ns.Container = { services: services, providerFor: providerFor };
})(globalThis.BankSync);
