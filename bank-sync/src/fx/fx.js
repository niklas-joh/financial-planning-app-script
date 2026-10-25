/**
 * @fileoverview Pure FX logic: rate lookup with previous-business-day fallback,
 * conversion to the reporting currency, and which date ranges still need rates.
 * Rates are ECB reference rates quoted as units of a currency per 1 EUR, so
 * `amountEur = amount / rate`. No Apps Script services here.
 */
(function (ns) {
  const DAY_MS = 86400000;

  /**
   * @typedef {Object} Rate
   * @property {string} date - ISO date the rate was published for.
   * @property {number} rate - Units of the currency per 1 unit of the reporting currency.
   */

  /**
   * Rates per currency, each sorted by date ascending, one entry per date.
   * @typedef {Object<string, Array<Rate>>} RateTable
   */

  /**
   * @typedef {Object} Conversion
   * @property {number} amountEur - Amount in the reporting currency, rounded to cents.
   * @property {number} fxRate - Rate used (1 for the reporting currency itself).
   * @property {string} fxRateDate - Date of the rate used.
   */

  /**
   * Whole days from one ISO date to another.
   * @param {string} from - YYYY-MM-DD.
   * @param {string} to - YYYY-MM-DD.
   * @returns {number} Days (negative if `to` is earlier).
   */
  function daysBetween(from, to) {
    return Math.round((Date.parse(to + 'T00:00:00Z') - Date.parse(from + 'T00:00:00Z')) / DAY_MS);
  }

  /**
   * The rate that applies on a date: that day's rate, else the latest earlier one
   * (weekends and ECB holidays have no rate), but never one more than `maxAgeDays` older.
   * @param {Array<Rate>} series - Rates sorted by date ascending.
   * @param {string} date - ISO date to convert at.
   * @param {number} maxAgeDays - Largest allowed gap between the rate's date and `date`.
   * @returns {?Rate} The rate, or null if none is close enough.
   */
  function rateOn(series, date, maxAgeDays) {
    let lo = 0;
    let hi = (series || []).length - 1;
    let found = null;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (series[mid].date <= date) {
        found = series[mid];
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return found && daysBetween(found.date, date) <= maxAgeDays ? found : null;
  }

  /**
   * Converts an amount at a rate quoted per unit of the reporting currency.
   * @param {number} amount - Amount in the original currency.
   * @param {number} rate - Units of that currency per 1 unit of the reporting currency.
   * @returns {number} Amount in the reporting currency, rounded to cents (never -0).
   */
  function toReporting(amount, rate) {
    return Math.round((amount / rate) * 100) / 100 + 0;
  }

  /**
   * Converts one amount on one date.
   * @param {number} amount - Original amount (signed).
   * @param {string} currency - ISO 4217 code of `amount`.
   * @param {string} date - ISO date to convert at.
   * @param {RateTable} table - Known rates.
   * @param {{reporting: string, maxAgeDays: number}} opts - Reporting currency and fallback limit.
   * @returns {?Conversion} The conversion, or null if no usable rate is known.
   */
  function convert(amount, currency, date, table, opts) {
    if (typeof amount !== 'number' || isNaN(amount) || !currency || !date) return null;
    if (currency === opts.reporting) return { amountEur: amount, fxRate: 1, fxRateDate: date };
    const rate = rateOn(table[currency], date, opts.maxAgeDays);
    return rate ? { amountEur: toReporting(amount, rate.rate), fxRate: rate.rate, fxRateDate: rate.date } : null;
  }

  /**
   * Adds rates to a series; a new rate for a date already present replaces it.
   * @param {Array<Rate>} series - Existing rates (not mutated).
   * @param {Array<Rate>} incoming - New rates.
   * @returns {Array<Rate>} Combined rates sorted by date ascending.
   */
  function mergeSeries(series, incoming) {
    const byDate = {};
    (series || []).concat(incoming || []).forEach(function (r) { byDate[r.date] = r; });
    return Object.keys(byDate).sort().map(function (date) { return byDate[date]; });
  }

  /**
   * Date ranges to fetch so `series` covers `from`..`to`. Only the edges are fetched, so
   * the stored series stays one contiguous block (a date inside it with no rate is an
   * ECB holiday, not a gap). A range starting on a non-business day is answered from the
   * previous business day, which gives the fallback for the first dates too.
   * @param {Array<Rate>} series - Known rates sorted ascending.
   * @param {string} from - First date needed.
   * @param {string} to - Last date needed (usually today).
   * @returns {Array<{from: string, to: string}>} At most two ranges, older first.
   */
  function missingRanges(series, from, to) {
    if (!series || !series.length) return [{ from: from, to: to }];
    const first = series[0].date;
    const last = series[series.length - 1].date;
    const ranges = [];
    if (from < first) ranges.push({ from: from, to: first });
    if (to > last) ranges.push({ from: last, to: to });
    return ranges;
  }

  /**
   * Whether a row still needs an EUR amount. Rows that have one are never recalculated.
   * @param {{amountEur: *}} row - Stored row.
   * @returns {boolean} True if `amountEur` is empty.
   */
  function needsConversion(row) {
    return row.amountEur === null || row.amountEur === undefined || row.amountEur === '';
  }

  /**
   * Whether a conversion can be stored for good. A rate from an earlier day is final for a
   * past date (that day had no ECB rate), but not for today or later: today's rate may
   * simply not be published yet (around 16:00 CET).
   * @param {Conversion} c - Conversion found.
   * @param {string} date - Date the row converts at.
   * @param {string} today - ISO date of today.
   * @returns {boolean} True if it should be stored now.
   */
  function isFinal(c, date, today) {
    return c.fxRateDate === date || date < today;
  }

  ns.Fx = {
    needsConversion: needsConversion,
    isFinal: isFinal,
    rateOn: rateOn,
    toReporting: toReporting,
    convert: convert,
    mergeSeries: mergeSeries,
    missingRanges: missingRanges,
    daysBetween: daysBetween
  };
})(globalThis.BankSync);
