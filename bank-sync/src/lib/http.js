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
   * Builds the secret-free error for a failed response.
   * @param {number} status - HTTP status.
   * @param {*} body - Parsed body (may carry `message` and `error`, the API's error code).
   * @returns {Error} HttpError.
   */
  function toError(status, body) {
    const detail = body && body.message ? String(body.message).slice(0, 300) : 'request failed';
    return ns.Errors.http(status, ns.Redact.text(detail), body && typeof body.error === 'string' ? body.error : '');
  }

  /**
   * Creates a client.
   * @param {HttpDeps} deps - Injected fetch/sleep.
   * @returns {{request: function(HttpRequest): *}} Client whose `request` returns parsed JSON.
   * @throws {Error} HttpError from `request` on non-2xx after retries.
   */
  function create(deps) {
    const maxRetries = deps.maxRetries === undefined ? 3 : deps.maxRetries;

    function request(req) {
      const params = { method: req.method, headers: req.headers || {}, muteHttpExceptions: true };
      if (req.body !== undefined) {
        params.contentType = 'application/json';
        params.payload = JSON.stringify(req.body);
      }
      for (let attempt = 0; ; attempt++) {
        const response = deps.fetch(req.url, params);
        const status = response.getResponseCode();
        const body = parseBody(response.getContentText());
        if (status >= 200 && status < 300) return body;
        if (!isRetryable(status) || attempt >= maxRetries) throw toError(status, body);
        const wait = deps.retryDelayMs
          ? deps.retryDelayMs(status, response.getAllHeaders ? response.getAllHeaders() : {}, attempt)
          : backoffMs(attempt);
        if (wait === null || wait === undefined) throw toError(status, body);
        deps.sleep(wait);
      }
    }

    return { request: request };
  }

  ns.Http = { create: create, backoffMs: backoffMs };
})(globalThis.BankSync);
