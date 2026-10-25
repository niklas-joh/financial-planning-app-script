/**
 * @fileoverview Pure retry decision for Trading 212's rate limits. Limits count per
 * account (whatever the key), so a call can hit one used up by an earlier run or another
 * app; the wait then comes from the `x-ratelimit-reset` header rather than a fixed backoff.
 * Plugged into `Http.create` as `retryDelayMs`.
 */
(function (ns) {
  const RESET_HEADER = 'x-ratelimit-reset';
  /** Margin after the reset time, so the retry does not land a moment too early. */
  const MARGIN_MS = 250;

  /**
   * Reads a header regardless of its case.
   * @param {Object<string,*>} headers - Response headers.
   * @param {string} name - Lowercase header name.
   * @returns {*} The value, or undefined.
   */
  function header(headers, name) {
    const key = Object.keys(headers || {}).filter(function (k) { return k.toLowerCase() === name; })[0];
    return key === undefined ? undefined : headers[key];
  }

  /**
   * Epoch ms of the reset. The docs say "Unix timestamp" without the unit, so both seconds and
   * milliseconds are accepted.
   * @param {*} value - Header value.
   * @returns {?number} Epoch ms, or null if absent or invalid.
   */
  function resetMs(value) {
    const n = Number(Array.isArray(value) ? value[0] : value);
    if (value === undefined || value === null || value === '' || !isFinite(n) || n <= 0) return null;
    return n < 1e12 ? n * 1000 : n;
  }

  /**
   * How long to wait before retrying.
   * @param {number} status - HTTP status (429 or 5xx).
   * @param {Object<string,*>} headers - Response headers.
   * @param {number} attempt - 0-based retry number.
   * @param {number} nowMs - Current epoch ms.
   * @param {number} msLeft - Time left in the run's budget.
   * @returns {?number} Wait in ms: until the reset for a 429 that names one, else the standard backoff;
   *   null (give up) when the wait would exceed `msLeft` or `Config.TRADING212.MAX_RETRY_WAIT_MS`.
   */
  function delayMs(status, headers, attempt, nowMs, msLeft) {
    const reset = status === 429 ? resetMs(header(headers, RESET_HEADER)) : null;
    const wait = reset === null ? ns.Http.backoffMs(attempt) : Math.max(0, reset - nowMs) + MARGIN_MS;
    return wait > ns.Config.TRADING212.MAX_RETRY_WAIT_MS || wait > msLeft ? null : wait;
  }

  /**
   * Builds the `retryDelayMs` hook for `Http.create`.
   * @param {function(): number} nowMs - Clock.
   * @param {function(): number} [msLeft] - Time left in the run; unlimited if omitted.
   * @returns {function(number, Object, number): ?number} Hook.
   */
  function hook(nowMs, msLeft) {
    return function (status, headers, attempt) {
      return delayMs(status, headers, attempt, nowMs(), msLeft ? msLeft() : Infinity);
    };
  }

  ns.T212RateLimit = { delayMs: delayMs, hook: hook };
})(globalThis.BankSync);
