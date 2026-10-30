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
   * Masks an IBAN down to its last four characters. Always masks; stored data goes through
   * `maskAccount`, which applies the `MASK_ACCOUNT_NUMBERS` setting.
   * @param {?string} iban - Full IBAN, or empty.
   * @returns {string} e.g. '****1234', or '' for empty input.
   */
  function maskIban(iban) {
    return iban ? '****' + String(iban).replace(/\s/g, '').slice(-4) : '';
  }

  /**
   * The one place an account, card or reference number is masked before it is stored. Follows
   * `Config.MASK_ACCOUNT_NUMBERS`: when on, the value is cut to its last four characters (`maskIban`);
   * when off, it is returned unchanged. Usable as a `String.replace` callback (extra arguments are ignored).
   * @param {*} value - Number as found (IBAN, card number, end-to-end id...), or empty.
   * @returns {string} The masked or the unchanged value as text; '' for null, undefined or ''.
   */
  function maskAccount(value) {
    if (value === null || value === undefined || value === '') return '';
    return ns.Config.MASK_ACCOUNT_NUMBERS ? maskIban(value) : String(value);
  }

  /** Masked values as `maskIban` writes them, e.g. '****1234', anywhere in a text. */
  const MASKED = /\*{4}[^\s*]{1,4}/;

  /**
   * Whether a stored value still holds a masked number (written while masking was on).
   * @param {*} value - Cell value.
   * @returns {boolean} True for text containing e.g. '****1234'.
   */
  function isMasked(value) {
    return typeof value === 'string' && MASKED.test(value);
  }

  ns.Redact = { text: text, maskIban: maskIban, maskAccount: maskAccount, isMasked: isMasked };
})(globalThis.BankSync);
