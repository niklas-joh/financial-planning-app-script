/**
 * @fileoverview Frankfurter v2 client for ECB reference rates (https://frankfurter.dev).
 * Verified 2026-10-25: `GET /v2/rates?from&to&base&quotes&providers=ecb` returns a JSON array of
 * `{date, base, quote, rate}`, one entry per ECB business day; `from` on a weekend or holiday
 * starts at the previous business day; `to` past the latest publication is clipped; a range
 * with no data returns `[]`; an unknown currency is HTTP 422. Without `providers=ecb` the API
 * blends other sources and also returns weekend dates, so the parameter is required.
 */
(function (ns) {
  const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

  /**
   * Checks one entry and turns it into a Rate.
   * @param {*} entry - Array element from the response.
   * @param {string} base - Expected base currency.
   * @param {string} quote - Expected quote currency.
   * @returns {import('./fx.js').Rate} Rate.
   * @throws {Error} DataError if the entry is not a usable rate.
   */
  function toRate(entry, base, quote) {
    const ok = entry && entry.base === base && entry.quote === quote && ISO_DATE.test(entry.date) &&
      typeof entry.rate === 'number' && isFinite(entry.rate) && entry.rate > 0;
    if (!ok) throw ns.Errors.create('DataError', 'Unexpected rate from Frankfurter for ' + quote);
    return { date: entry.date, rate: entry.rate };
  }

  /**
   * Creates a client.
   * @param {{request: function(Object): *}} http - JSON HTTP client.
   * @returns {{fetchRates: function(string, string, string): Array<import('./fx.js').Rate>}} Client.
   */
  function create(http) {
    const cfg = ns.Config.FX;
    const base = ns.Config.REPORTING_CURRENCY;
    return {
      /**
       * ECB rates for one currency over a date range, one call.
       * @param {string} currency - ISO 4217 quote currency.
       * @param {string} from - First date, inclusive (YYYY-MM-DD).
       * @param {string} to - Last date, inclusive (YYYY-MM-DD).
       * @returns {Array<import('./fx.js').Rate>} Rates (units of `currency` per 1 EUR), ascending.
       * @throws {Error} HttpError (e.g. 422 for an unknown currency) or DataError.
       */
      fetchRates: function (currency, from, to) {
        const params = { from: from, to: to, base: base, quotes: currency, providers: cfg.RATE_PROVIDER };
        const query = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
        const body = http.request({ method: 'GET', url: cfg.BASE_URL + '/rates?' + query, headers: { Accept: 'application/json' } });
        if (!Array.isArray(body)) throw ns.Errors.create('DataError', 'Unexpected response from Frankfurter for ' + currency);
        return body.map(function (entry) { return toRate(entry, base, currency); })
          .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
      }
    };
  }

  ns.Frankfurter = { create: create };
})(globalThis.BankSync);
