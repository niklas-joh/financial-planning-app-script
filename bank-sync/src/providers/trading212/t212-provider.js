/**
 * @fileoverview Trading212Provider: implements the BankProvider contract for one
 * Trading 212 account (one API key). Adds the optional `getHoldings`, and reads its
 * cursor-paged history in one window (`maxWindowDays: Infinity`).
 */
(function (ns) {
  /**
   * @typedef {Object} T212ProviderDeps
   * @property {import('../../core/config.js').ConnectionConfig} connection - Connection (holder) served.
   * @property {Object} client - T212Client.
   * @property {function(): number} nowMs - Clock.
   * @property {function(string): string} toLocalDate - ISO date-time to YYYY-MM-DD in the Bank Sync time zone.
   */

  /**
   * Memoizes a function of no arguments for the life of the provider (one sync run): the summary
   * feeds the account, the balances and the holdings' currency, and the API allows 1 call per 5 s.
   * @param {function(): *} fn - Function.
   * @returns {function(): *} Memoized function.
   */
  function once(fn) {
    let done = false;
    let value;
    return function () {
      if (!done) {
        value = fn();
        done = true;
      }
      return value;
    };
  }

  /**
   * Reads newest-first pages until one reaches back past `dateFrom`, the last page, or the page cap.
   * @param {function(?string): {items: Array<Object>, nextPagePath: ?string}} readPage - Page reader.
   * @param {function(Object): string} dateOf - Local date of an item.
   * @param {string} dateFrom - Oldest date needed.
   * @returns {Array<Object>} Items read (may include some older than `dateFrom`).
   * @throws {Error} DataError if the cap is reached before `dateFrom`.
   */
  function readBack(readPage, dateOf, dateFrom) {
    const maxPages = ns.Config.TRANSACTIONS.MAX_PAGES;
    let items = [];
    let path = null;
    for (let pages = 0; pages < maxPages; pages++) {
      const page = readPage(path);
      items = items.concat(page.items);
      const reachedStart = page.items.some(function (item) { return dateOf(item) < dateFrom; });
      if (!page.nextPagePath || reachedStart) return items;
      path = page.nextPagePath;
    }
    throw ns.Errors.create('DataError', 'Trading 212 history did not reach ' + dateFrom + ' within ' + maxPages + ' pages');
  }

  /**
   * Creates a provider for one connection.
   * @param {T212ProviderDeps} deps - Injected services.
   * @returns {import('../bank-provider.js').BankProvider} The provider.
   */
  function create(deps) {
    const summary = once(deps.client.summary);
    const account = once(function () { return ns.T212Normalize.account(summary(), deps.connection); });
    const isoNow = function () { return new Date(deps.nowMs()).toISOString(); };
    const toLocalDate = deps.toLocalDate;
    /** Local date for the paging stop; an invalid date never stops it (the normalizer rejects the item). */
    const pageDate = function (value) { return value && !isNaN(Date.parse(value)) ? toLocalDate(value) : '\uffff'; };

    return ns.BankProvider.assertBankProvider({
      maxWindowDays: Infinity,

      authorize: function () {
        throw ns.Errors.config('Trading 212 uses an API key: set ' + ns.Config.PROPERTY_KEYS.T212_API_KEY_PREFIX + '<HOLDER> and ' +
          ns.Config.PROPERTY_KEYS.T212_API_SECRET_PREFIX + '<HOLDER> in Script Properties');
      },

      listAccounts: function () {
        return [account()];
      },

      getBalances: function (accountId) {
        return ns.T212Normalize.balances(summary(), accountId, isoNow());
      },

      getHoldings: function (accountId) {
        return ns.T212Normalize.holdings(deps.client.positions(), accountId, account().currency, isoNow());
      },

      /**
       * Cash movements and dividends booked in a date range (local dates, inclusive).
       * @param {string} accountId - Account id.
       * @param {{dateFrom: string, dateTo: string}} range - ISO dates.
       * @returns {Array<import('../bank-provider.js').Transaction>} Transactions in the range.
       * @throws {Error} DataError if the history cannot be read back to `dateFrom`.
       */
      getTransactions: function (accountId, range) {
        const cash = readBack(deps.client.transactionsPage, function (i) { return pageDate(i.dateTime); }, range.dateFrom)
          .map(function (i) { return ns.T212Normalize.cashMovement(i, accountId, toLocalDate); });
        const dividends = readBack(deps.client.dividendsPage, function (i) { return pageDate(i.paidOn); }, range.dateFrom)
          .map(function (i) { return ns.T212Normalize.dividend(i, accountId, toLocalDate); });
        return cash.concat(dividends).filter(function (t) {
          return t.bookingDate >= range.dateFrom && t.bookingDate <= range.dateTo;
        });
      }
    });
  }

  ns.Trading212Provider = { create: create };
})(globalThis.BankSync);
