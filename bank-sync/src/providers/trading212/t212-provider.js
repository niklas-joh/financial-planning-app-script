/**
 * @fileoverview Trading212Provider: implements the BankProvider contract for one
 * Trading 212 account (one API key). Adds the optional `getHoldings`, reads its
 * cursor-paged history in one window (`maxWindowDays: Infinity`), and on top of that
 * backfills the full history once, a bounded number of pages per run (`T212History`).
 */
(function (ns) {
  /**
   * @typedef {Object} T212ProviderDeps
   * @property {import('../../core/config.js').ConnectionConfig} connection - Connection (holder) served.
   * @property {Object} client - T212Client.
   * @property {function(): number} nowMs - Clock.
   * @property {function(string): string} toLocalDate - ISO date-time to YYYY-MM-DD in the Bank Sync time zone.
   * @property {{load: function(): Object, save: function(Object): void}} [history] - Backfill state
   *   (`T212History.store`); without it only the incremental history is read.
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
   * @param {Array<import('../bank-provider.js').Transaction>} rows - Transactions.
   * @returns {Array<import('../bank-provider.js').Transaction>} The first row of each transaction id.
   */
  function uniqueById(rows) {
    const seen = {};
    return rows.filter(function (t) {
      if (seen[t.transactionId]) return false;
      seen[t.transactionId] = true;
      return true;
    });
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
    const feeds = {
      transactions: { readPage: deps.client.transactionsPage, normalize: ns.T212Normalize.cashMovement },
      dividends: { readPage: deps.client.dividendsPage, normalize: ns.T212Normalize.dividend }
    };

    /**
     * Continues the one-time backfill of every feed not done yet. The position of a feed only moves
     * once its items are normalized, so an item that fails is never skipped; progress is saved also
     * when a later feed fails.
     * @param {string} accountId - Account id.
     * @returns {Array<import('../bank-provider.js').Transaction>} Older transactions read in this run.
     */
    function backfill(accountId) {
      if (!deps.history) return [];
      const state = deps.history.load();
      let rows = [];
      try {
        ns.T212History.FEEDS.forEach(function (name) {
          if (state[name].done) return;
          const read = ns.T212History.step(feeds[name].readPage, state[name], ns.Config.TRADING212.BACKFILL_PAGES_PER_RUN);
          rows = rows.concat(read.items.map(function (i) { return feeds[name].normalize(i, accountId, toLocalDate); }));
          state[name] = read.feed;
        });
      } finally {
        deps.history.save(state);
      }
      return rows;
    }

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
       * Cash movements and dividends booked in a date range (local dates, inclusive), plus whatever
       * older history the one-time backfill reads in this run (never later than `dateTo`).
       * @param {string} accountId - Account id.
       * @param {{dateFrom: string, dateTo: string}} range - ISO dates.
       * @returns {Array<import('../bank-provider.js').Transaction>} Transactions, one per id.
       * @throws {Error} DataError if the history cannot be read back to `dateFrom`.
       */
      getTransactions: function (accountId, range) {
        const cash = readBack(deps.client.transactionsPage, function (i) { return pageDate(i.dateTime); }, range.dateFrom)
          .map(function (i) { return ns.T212Normalize.cashMovement(i, accountId, toLocalDate); });
        const dividends = readBack(deps.client.dividendsPage, function (i) { return pageDate(i.paidOn); }, range.dateFrom)
          .map(function (i) { return ns.T212Normalize.dividend(i, accountId, toLocalDate); });
        const recent = cash.concat(dividends).filter(function (t) {
          return t.bookingDate >= range.dateFrom && t.bookingDate <= range.dateTo;
        });
        const older = backfill(accountId).filter(function (t) { return t.bookingDate <= range.dateTo; });
        return uniqueById(recent.concat(older));
      }
    });
  }

  ns.Trading212Provider = { create: create };
})(globalThis.BankSync);
