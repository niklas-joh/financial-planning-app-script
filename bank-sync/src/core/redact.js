/**
 * @fileoverview Scrubs secrets from any text before it is logged or shown.
 */
(function (ns) {
  const PATTERNS = [
    [/-----BEGIN [A-Z ]+-----[\s\S]*?-----END [A-Z ]+-----/g, '[REDACTED:pem]'],
    [/eyJ[\w-]+\.[\w-]+\.[\w-]+/g, '[REDACTED:jwt]'],
    [/Bearer\s+[\w.~+/-]+=*/gi, 'Bearer [REDACTED]'],
    [/Basic\s+[A-Za-z0-9+/]{16,}=*/g, 'Basic [REDACTED]'],
    [/\b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b/g, '[REDACTED:iban]']
  ];

  /**
   * Replaces PEM blocks, JWTs, bearer tokens and IBANs in a string.
   * @param {*} value - Anything; coerced to string.
   * @returns {string} Text with secrets replaced by placeholders.
   */
  function text(value) {
    return PATTERNS.reduce(function (out, pair) {
      return out.replace(pair[0], pair[1]);
    }, String(value));
  }

  /**
   * Masks an IBAN down to its last four characters.
   * @param {?string} iban - Full IBAN, or empty.
   * @returns {string} e.g. '****1234', or '' for empty input.
   */
  function maskIban(iban) {
    return iban ? '****' + String(iban).replace(/\s/g, '').slice(-4) : '';
  }

  ns.Redact = { text: text, maskIban: maskIban };
})(globalThis.BankSync);
