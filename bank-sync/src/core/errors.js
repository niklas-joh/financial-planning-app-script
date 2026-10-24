/**
 * @fileoverview Small error factory so callers can branch on `error.name`.
 * Messages must never contain tokens, keys or IBANs.
 */
(function (ns) {
  /**
   * Creates an Error with a custom name and optional extra fields.
   * @param {string} name - Error name, e.g. 'HttpError'.
   * @param {string} message - Human readable, secret-free message.
   * @param {Object} [extra] - Extra properties copied onto the error.
   * @returns {Error} The new error.
   */
  function create(name, message, extra) {
    const error = new Error(message);
    error.name = name;
    Object.assign(error, extra || {});
    return error;
  }

  ns.Errors = {
    create: create,

    /**
     * @param {number} status - HTTP status code.
     * @param {string} detail - Redacted API message.
     * @returns {Error} HttpError carrying `status`.
     */
    http: function (status, detail) {
      return create('HttpError', 'HTTP ' + status + ': ' + detail, { status: status });
    },

    /**
     * @param {string} message - What is misconfigured (names only, no values).
     * @returns {Error} ConfigError.
     */
    config: function (message) {
      return create('ConfigError', message);
    },

    /**
     * @param {string} message - Why authorization is missing or invalid.
     * @returns {Error} AuthError.
     */
    auth: function (message) {
      return create('AuthError', message);
    },

    /**
     * @param {string} what - The unimplemented capability.
     * @returns {Error} NotImplementedError.
     */
    notImplemented: function (what) {
      return create('NotImplementedError', what + ' is not implemented yet');
    }
  };
})(globalThis.BankSync);
