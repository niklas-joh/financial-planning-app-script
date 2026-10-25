/**
 * @fileoverview Pure ISO-date (YYYY-MM-DD) arithmetic and fetch-window planning.
 * Dates are treated as UTC calendar days, which is what Enable Banking assumes.
 */
(function (ns) {
  const DAY_MS = 86400000;

  /**
   * @typedef {Object} DateWindow
   * @property {string} from - First day, inclusive (ISO date).
   * @property {string} to - Last day, inclusive (ISO date).
   */

  /**
   * Formats an epoch time as an ISO calendar date (UTC).
   * @param {number} ms - Epoch milliseconds.
   * @returns {string} YYYY-MM-DD.
   */
  function isoDate(ms) {
    return new Date(ms).toISOString().slice(0, 10);
  }

  /**
   * Adds whole days to an ISO date.
   * @param {string} iso - YYYY-MM-DD.
   * @param {number} days - Days to add (may be negative).
   * @returns {string} YYYY-MM-DD.
   */
  function addDays(iso, days) {
    return isoDate(Date.parse(iso + 'T00:00:00Z') + days * DAY_MS);
  }

  /**
   * Chooses the first day to fetch for one account.
   * @param {string} latestBooked - Latest stored booked date for the account, or ''.
   * @param {string} today - ISO date of today.
   * @param {{INITIAL_LOOKBACK_DAYS: number, OVERLAP_DAYS: number}} cfg - Tuning.
   * @returns {string} ISO date: lookback on first run, else the latest booked date
   *   minus the overlap (so late-booked and pending items are re-read).
   */
  function startFor(latestBooked, today, cfg) {
    return latestBooked
      ? addDays(latestBooked, -cfg.OVERLAP_DAYS)
      : addDays(today, -cfg.INITIAL_LOOKBACK_DAYS);
  }

  /**
   * Splits a range into consecutive, non-overlapping windows, oldest first.
   * @param {string} from - First day, inclusive.
   * @param {string} to - Last day, inclusive.
   * @param {number} windowDays - Maximum days per window (>= 1); Infinity gives one window for the whole range.
   * @returns {Array<DateWindow>} Windows; empty when from is after to.
   */
  function planWindows(from, to, windowDays) {
    if (!isFinite(windowDays)) return from <= to ? [{ from: from, to: to }] : [];
    const windows = [];
    for (let start = from; start <= to; start = addDays(start, windowDays)) {
      const end = addDays(start, windowDays - 1);
      windows.push({ from: start, to: end < to ? end : to });
    }
    return windows;
  }

  ns.DateRange = { isoDate: isoDate, addDays: addDays, startFor: startFor, planWindows: planWindows };
})(globalThis.BankSync);
