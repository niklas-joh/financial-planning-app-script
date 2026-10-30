/**
 * @fileoverview JSON HTTP client with retry/backoff. `fetch` and `sleep` are
 * injected (UrlFetchApp.fetch and Utilities.sleep in Apps Script).
 */
(function (ns) {
  const RETRY_BASE_MS = 500;

  /**
   * @typedef {Object} HttpRequest
   * @property {string} method - HTTP method.
   * @property {string} url - Absolute URL.
   * @property {Object<string,string>} [headers] - Request headers.
   * @property {Object} [body] - JSON-serializable body.
   * @property {boolean} [text] - Return a successful body as text instead of parsed JSON (e.g. a CSV file).
   * @property {boolean} [followRedirects] - false: a redirect is not followed but returned as `{redirect: <Location>}`,
   *   so the caller decides where its headers may go (UrlFetchApp follows redirects with the same headers otherwise).
   */

  /**
   * @typedef {Object} HttpDeps
   * @property {function(string, Object): {getResponseCode: function(): number, getContentText: function(): string}} fetch
   * @property {function(number): void} sleep
   * @property {number} [maxRetries] - Retries on 429/5xx (default 3).
   * @property {function(number, Object<string,string>, number): ?number} [retryDelayMs] - Decides the wait
   *   before retry `attempt` (0-based) after a 429/5xx from the status and the response headers; null gives up
   *   and throws the error. Without it the wait is 500 ms doubled per attempt.
   */

  /**
   * Whether a status is worth retrying.
   * @param {number} status - HTTP status.
   * @returns {boolean} True for 429 and 5xx.
   */
  function isRetryable(status) {
    return status === 429 || status >= 500;
  }

  /**
   * Default wait before a retry.
   * @param {number} attempt - 0-based retry number.
   * @returns {number} 500 ms doubled per attempt.
   */
  function backoffMs(attempt) {
    return RETRY_BASE_MS * Math.pow(2, attempt);
  }

  /**
   * Parses a response body as JSON, tolerating empty/non-JSON bodies.
   * @param {string} text - Response body.
   * @returns {*} Parsed value, or null when empty/unparseable.
   */
  function parseBody(text) {
    try {
      return text ? JSON.parse(text) : null;
    } catch (e) {
      return null;
    }
  }

  /**
   * Builds the secret-free error for a failed response. The detail is the body's `message`; without one, the
   * body itself, shortened (some APIs, e.g. Trading 212's 400, name the problem in other fields), so the
   * error says what was wrong.
   * @param {number} status - HTTP status.
   * @param {*} body - Parsed body (may carry `message` and `error`, the API's error code).
   * @param {string} text - Raw body.
   * @returns {Error} HttpError.
   */
  function toError(status, body, text) {
    const raw = body && body.message ? String(body.message)
      : body && typeof body === 'object' ? JSON.stringify(body) : String(text || '').trim();
    const detail = raw ? raw.slice(0, 300) : 'request failed';
    return ns.Errors.http(status, ns.Redact.text(detail), body && typeof body.error === 'string' ? body.error : '');
  }

  /**
   * @param {Object<string,*>} headers - Response headers.
   * @param {string} name - Header name, any case.
   * @returns {string} Its value, or ''.
   */
  function header(headers, name) {
    const key = Object.keys(headers || {}).filter(function (k) { return k.toLowerCase() === name.toLowerCase(); })[0];
    return key ? String(headers[key]) : '';
  }

  /**
   * Creates a client.
   * @param {HttpDeps} deps - Injected fetch/sleep.
   * @returns {{request: function(HttpRequest): *}} Client whose `request` returns parsed JSON (text with `text`).
   * @throws {Error} HttpError from `request` on non-2xx after retries.
   */
  function create(deps) {
    const maxRetries = deps.maxRetries === undefined ? 3 : deps.maxRetries;

    function request(req) {
      const params = { method: req.method, headers: req.headers || {}, muteHttpExceptions: true };
      if (req.followRedirects === false) params.followRedirects = false;
      if (req.body !== undefined) {
        params.contentType = 'application/json';
        params.payload = JSON.stringify(req.body);
      }
      for (let attempt = 0; ; attempt++) {
        const response = deps.fetch(req.url, params);
        const status = response.getResponseCode();
        if (status >= 200 && status < 300 && req.text) return response.getContentText();
        if (req.followRedirects === false && status >= 300 && status < 400) {
          return { redirect: header(response.getAllHeaders ? response.getAllHeaders() : {}, 'Location') };
        }
        const text = response.getContentText();
        const body = parseBody(text);
        if (status >= 200 && status < 300) return body;
        if (!isRetryable(status) || attempt >= maxRetries) throw toError(status, body, text);
        const wait = deps.retryDelayMs
          ? deps.retryDelayMs(status, response.getAllHeaders ? response.getAllHeaders() : {}, attempt)
          : backoffMs(attempt);
        if (wait === null || wait === undefined) throw toError(status, body, text);
        deps.sleep(wait);
      }
    }

    return { request: request };
  }

  ns.Http = { create: create, backoffMs: backoffMs };
})(globalThis.BankSync);
