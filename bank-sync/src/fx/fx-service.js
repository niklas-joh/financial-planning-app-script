/**
 * @fileoverview Injectable FX service: makes sure the rates a run needs are known
 * (Rates tab, CacheService, then the rates source for the missing edges only), then
 * converts transactions at their booking-date rate and balances at the latest rate.
 * The rates source, the tab store and the cache are passed in, so this runs in Node.
 */
(function (ns) {
  const CACHE_KEY = 'fx:rates';
  const CHECKED_PREFIX = 'fx:checked:';

  /**
   * @typedef {Object} FxServiceDeps
   * @property {function(string, string, string): Array<import('./fx.js').Rate>} fetchRates - Rates for
   *   (currency, from, to), both dates inclusive; one HTTP call.
   * @property {{load: function(): import('./fx.js').RateTable, save: function(import('./fx.js').RateTable): number}} rates
   *   The Rates tab.
   * @property {{get: function(string): ?string, put: function(string, string, number): void, remove: function(string): void}} cache
   *   Script cache (CacheService-like).
   * @property {function(): string} todayIso - Today's date in the configured time zone.
   */

  /**
   * @typedef {Object} FxResult
   * @property {number} changed - Rows whose EUR columns were set or corrected.
   * @property {number} missing - Rows still without a rate (unknown currency or no rate close enough).
   * @property {number} waiting - Rows dated today whose rate is not published yet; converted on a later run.
   * @property {Array<string>} failed - Redacted "CUR: error" for each currency whose fetch failed.
   */

  /**
   * Whether two rows carry the same EUR columns.
   * @param {Object} row - Stored row.
   * @param {import('./fx.js').Conversion} c - Fresh conversion.
   * @returns {boolean} True if nothing would change.
   */
  function sameConversion(row, c) {
    return row.amountEur === c.amountEur && row.fxRate === c.fxRate && row.fxRateDate === c.fxRateDate;
  }

  /**
   * Creates the service. Rates are loaded once and fetched at most once per missing range
   * per currency for the life of the instance (one sync run).
   * @param {FxServiceDeps} deps - Injected source, store, cache and clock.
   * @returns {{convertTransactions: function(Array<Object>): {rows: Array<Object>} & FxResult,
   *   convertBalances: function(Array<Object>): {rows: Array<Object>} & FxResult}} Service.
   */
  function create(deps) {
    const cfg = ns.Config.FX;
    const opts = { reporting: ns.Config.REPORTING_CURRENCY, maxAgeDays: cfg.MAX_FALLBACK_DAYS };
    const failed = {};
    let table = null;

    /** @returns {import('./fx.js').RateTable} Known rates, from the cache or else the tab. */
    function load() {
      if (table) return table;
      const cached = deps.cache.get(CACHE_KEY);
      if (cached) {
        try {
          table = JSON.parse(cached);
          return table;
        } catch (e) {
          deps.cache.remove(CACHE_KEY);
        }
      }
      table = deps.rates.load();
      return table;
    }

    /** Writes the rates to the tab and refreshes the cache (dropped if too large). */
    function persist() {
      deps.rates.save(table);
      const json = JSON.stringify(table);
      if (json.length <= cfg.CACHE_MAX_CHARS) deps.cache.put(CACHE_KEY, json, cfg.CACHE_SECONDS);
      else deps.cache.remove(CACHE_KEY);
    }

    /**
     * Fetches whatever is missing so each currency is covered from its earliest needed date to today.
     * A range already asked for in the last `CACHE_SECONDS` is not asked again (today's rate is
     * published in the afternoon, and older history may simply not exist).
     * @param {Object<string,string>} earliest - Earliest needed date per currency.
     */
    function ensure(earliest) {
      const today = deps.todayIso();
      const checked = [];
      load();
      Object.keys(earliest).forEach(function (currency) {
        if (currency === opts.reporting || failed[currency]) return;
        ns.Fx.missingRanges(table[currency], earliest[currency], today).forEach(function (range) {
          const marker = CHECKED_PREFIX + currency + ':' + range.from + ':' + range.to;
          if (failed[currency] || deps.cache.get(marker)) return;
          try {
            table[currency] = ns.Fx.mergeSeries(table[currency], deps.fetchRates(currency, range.from, range.to));
            checked.push(marker);
          } catch (e) {
            failed[currency] = currency + ': ' + ns.Redact.text(e.name + ': ' + e.message);
          }
        });
      });
      if (!checked.length) return;
      persist();
      // Marked only once saved, so a failed save is retried on the next run.
      checked.forEach(function (marker) { deps.cache.put(marker, '1', cfg.CACHE_SECONDS); });
    }

    /**
     * Converts the rows that `needs` selects; every other row is returned as it is.
     * @param {Array<Object>} rows - Rows with `amount` and `currency`.
     * @param {function(Object): string} dateOf - Date each row converts at.
     * @param {function(Object): boolean} needs - Whether a row should be converted.
     * @param {function(import('./fx.js').Conversion, string): boolean} accept - Whether a conversion
     *   found for a row on that date may be stored now (false leaves the row for a later run).
     * @returns {{rows: Array<Object>} & FxResult} New rows (inputs are not mutated) and counts.
     */
    function apply(rows, dateOf, needs, accept) {
      const todo = rows.filter(needs);
      const earliest = {};
      todo.forEach(function (row) {
        const date = dateOf(row);
        if (row.currency && (!earliest[row.currency] || date < earliest[row.currency])) earliest[row.currency] = date;
      });
      if (todo.length) ensure(earliest);
      let changed = 0;
      let missing = 0;
      let waiting = 0;
      const out = rows.map(function (row) {
        if (!needs(row)) return row;
        const date = dateOf(row);
        const c = ns.Fx.convert(row.amount, row.currency, date, table, opts);
        if (!c) {
          missing++;
          return row;
        }
        if (!accept(c, date)) {
          waiting++;
          return row;
        }
        if (sameConversion(row, c)) return row;
        changed++;
        return Object.assign({}, row, c);
      });
      return {
        rows: out,
        changed: changed,
        missing: missing,
        waiting: waiting,
        failed: Object.keys(failed).map(function (k) { return failed[k]; })
      };
    }

    return {
      /**
       * Converts the transactions that have no EUR amount yet (new rows, or rows whose amount,
       * currency or date changed, which the merge clears) at the rate of their booking date
       * (else value date; undated pending rows at the latest rate), falling back to the previous
       * business day. Rows that already have an EUR amount are never recalculated. A row dated
       * today whose rate is not published yet is left empty for a later run rather than stored
       * with the previous day's rate.
       * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Stored transactions.
       * @returns {{rows: Array<import('../providers/bank-provider.js').Transaction>} & FxResult} Converted rows and counts.
       */
      convertTransactions: function (rows) {
        const today = deps.todayIso();
        return apply(rows,
          function (t) { return ns.TxnMerge.effectiveDate(t) || today; },
          ns.Fx.needsConversion,
          function (c, date) { return ns.Fx.isFinal(c, date, today); });
      },

      /**
       * Converts each balance at the latest known rate (balances are fresh on every sync).
       * @param {Array<import('../providers/bank-provider.js').Balance>} balances - Fresh balances.
       * @returns {{rows: Array<import('../providers/bank-provider.js').Balance>} & FxResult} Converted balances and counts.
       */
      convertBalances: function (balances) {
        const today = deps.todayIso();
        const always = function () { return true; };
        return apply(balances, function () { return today; }, always, always);
      }
    };
  }

  ns.FxService = { create: create };
})(globalThis.BankSync);
